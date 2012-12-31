VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmCMP001 
   BackColor       =   &H8000000A&
   Caption         =   "CMP001"
   ClientHeight    =   6075
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9945
   Icon            =   "CMP001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   9945
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "CMP001.frx":08CA
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCmpCode 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2010
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   5655
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   9735
      Begin TabDlg.SSTab tabDetailInfo 
         Height          =   3405
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6006
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Address"
         TabPicture(0)   =   "CMP001.frx":2FCD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtCmpRptEngAdd"
         Tab(0).Control(1)=   "txtCmpRptChiAdd"
         Tab(0).Control(2)=   "picCmpAdr"
         Tab(0).Control(3)=   "lblCmpRptEngAdd"
         Tab(0).Control(4)=   "lblCmpRptChiAdd"
         Tab(0).Control(5)=   "lblCmpAdr"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "General Info"
         TabPicture(1)   =   "CMP001.frx":2FE9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtCmpFax"
         Tab(1).Control(1)=   "txtCmpTel"
         Tab(1).Control(2)=   "txtCmpEmail"
         Tab(1).Control(3)=   "txtCmpWebSite"
         Tab(1).Control(4)=   "lblCmpFax"
         Tab(1).Control(5)=   "lblCmpTel"
         Tab(1).Control(6)=   "lblCmpEmail"
         Tab(1).Control(7)=   "lblCmpWebSite"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Accounting Info"
         TabPicture(2)   =   "CMP001.frx":3005
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "lblCmpPayCode"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lblCmpRetainAC"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "lblCmpExgMLCode"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "lblCmpTiMLCode"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lblCmpCurr"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "lblCmpSupMLCode"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "lblCmpExlMLCode"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "lblCmpTeMLCode"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "lblCmpSamMLCode"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "lblCmpDamMLCode"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "lblCmpAdjMLCode"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "lblCmpCurrEarn"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "cboCmpPayCode"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "cboCmpRetainAC"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "cboCmpExgMLCode"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "cboCmpTiMLCode"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "cboCmpCurr"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "cboCmpSupMLCode"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "cboCmpExlMLCode"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "cboCmpTeMLCode"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "cboCmpSamMLCode"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "cboCmpDamMLCode"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "cboCmpAdjMLCode"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "cboCmpCurrEarn"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).ControlCount=   24
         TabCaption(3)   =   "Bank Info"
         TabPicture(3)   =   "CMP001.frx":3021
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtCmpBankAC"
         Tab(3).Control(1)=   "txtCmpBankACName"
         Tab(3).Control(2)=   "lblCmpBankAC"
         Tab(3).Control(3)=   "lblCmpBankACName"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Remark"
         TabPicture(4)   =   "CMP001.frx":303D
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtCmpRemark"
         Tab(4).Control(1)=   "lblCmpRemark"
         Tab(4).ControlCount=   2
         Begin VB.ComboBox cboCmpCurrEarn 
            Height          =   300
            Left            =   7680
            TabIndex        =   16
            Top             =   720
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpAdjMLCode 
            Height          =   300
            Left            =   2760
            TabIndex        =   23
            Top             =   2160
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpDamMLCode 
            Height          =   300
            Left            =   2760
            TabIndex        =   21
            Top             =   1800
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpSamMLCode 
            Height          =   300
            Left            =   7680
            TabIndex        =   22
            Top             =   1800
            Width           =   1530
         End
         Begin VB.TextBox txtCmpRptEngAdd 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -72960
            TabIndex        =   7
            Top             =   1800
            Width           =   7275
         End
         Begin VB.TextBox txtCmpRptChiAdd 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   -72960
            TabIndex        =   8
            Top             =   2400
            Width           =   7275
         End
         Begin VB.TextBox txtCmpRemark 
            Enabled         =   0   'False
            Height          =   1740
            Left            =   -73800
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   7665
         End
         Begin VB.PictureBox picCmpAdr 
            BackColor       =   &H80000009&
            Height          =   1455
            Left            =   -72960
            ScaleHeight     =   1395
            ScaleWidth      =   7215
            TabIndex        =   54
            Top             =   240
            Width           =   7275
            Begin VB.TextBox txtCmpAddress 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   7100
            End
            Begin VB.TextBox txtCmpAddress 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   4
               Top             =   345
               Width           =   7100
            End
            Begin VB.TextBox txtCmpAddress 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   5
               Top             =   690
               Width           =   7100
            End
            Begin VB.TextBox txtCmpAddress 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   6
               Top             =   1035
               Width           =   7100
            End
         End
         Begin VB.TextBox txtCmpFax 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -69000
            TabIndex        =   10
            Top             =   360
            Width           =   2925
         End
         Begin VB.TextBox txtCmpTel 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   -73560
            TabIndex        =   9
            Top             =   360
            Width           =   2925
         End
         Begin VB.TextBox txtCmpEmail 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -73560
            TabIndex        =   11
            Top             =   720
            Width           =   7485
         End
         Begin VB.TextBox txtCmpWebSite 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   -73560
            TabIndex        =   12
            Top             =   1080
            Width           =   7485
         End
         Begin VB.ComboBox cboCmpTeMLCode 
            Height          =   300
            Left            =   7680
            TabIndex        =   20
            Top             =   1440
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpExlMLCode 
            Height          =   300
            Left            =   7680
            TabIndex        =   18
            Top             =   1080
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpSupMLCode 
            Height          =   300
            Left            =   7680
            TabIndex        =   24
            Top             =   2160
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpCurr 
            Height          =   300
            Left            =   7680
            TabIndex        =   14
            Top             =   360
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpTiMLCode 
            Height          =   300
            Left            =   2760
            TabIndex        =   19
            Top             =   1440
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpExgMLCode 
            Height          =   300
            Left            =   2760
            TabIndex        =   17
            Top             =   1080
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpRetainAC 
            Height          =   300
            Left            =   2760
            TabIndex        =   15
            Top             =   720
            Width           =   1530
         End
         Begin VB.ComboBox cboCmpPayCode 
            Height          =   300
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   1530
         End
         Begin VB.TextBox txtCmpBankAC 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -73560
            TabIndex        =   25
            Top             =   360
            Width           =   7485
         End
         Begin VB.TextBox txtCmpBankACName 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   -73560
            TabIndex        =   26
            Top             =   720
            Width           =   7485
         End
         Begin VB.Label lblCmpCurrEarn 
            Caption         =   "CMPRETAINAC"
            Height          =   240
            Left            =   4560
            TabIndex        =   62
            Top             =   795
            Width           =   2580
         End
         Begin VB.Label lblCmpAdjMLCode 
            Caption         =   "CMPADJMLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   61
            Top             =   2235
            Width           =   2580
         End
         Begin VB.Label lblCmpDamMLCode 
            Caption         =   "CMPDAMMLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   1875
            Width           =   2580
         End
         Begin VB.Label lblCmpSamMLCode 
            Caption         =   "CMPSAMMLCODE"
            Height          =   240
            Left            =   4560
            TabIndex        =   59
            Top             =   1875
            Width           =   2580
         End
         Begin VB.Label lblCmpRptEngAdd 
            Caption         =   "CMPRPTENGADD"
            Height          =   375
            Left            =   -74760
            TabIndex        =   58
            Top             =   1860
            Width           =   1740
         End
         Begin VB.Label lblCmpRptChiAdd 
            Caption         =   "CMPRPTCHIADD"
            Height          =   480
            Left            =   -74760
            TabIndex        =   57
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label lblCmpRemark 
            Caption         =   "CMPREMARK"
            Height          =   240
            Left            =   -74880
            TabIndex        =   56
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblCmpAdr 
            Caption         =   "CMPADR"
            Height          =   480
            Left            =   -74760
            TabIndex        =   55
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label lblCmpFax 
            Caption         =   "CMPFAX"
            Height          =   255
            Left            =   -70200
            TabIndex        =   53
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblCmpTel 
            Caption         =   "CMPTEL"
            Height          =   240
            Left            =   -74760
            TabIndex        =   52
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label lblCmpEmail 
            Caption         =   "CMPEMAIL"
            Height          =   255
            Left            =   -74760
            TabIndex        =   51
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label lblCmpWebSite 
            Caption         =   "CMPWEBSITE"
            Height          =   240
            Left            =   -74760
            TabIndex        =   50
            Top             =   1140
            Width           =   1380
         End
         Begin VB.Label lblCmpTeMLCode 
            Caption         =   "CMPTEMLCODE"
            Height          =   240
            Left            =   4560
            TabIndex        =   49
            Top             =   1515
            Width           =   2580
         End
         Begin VB.Label lblCmpExlMLCode 
            Caption         =   "CMPEXLMLCODE"
            Height          =   240
            Left            =   4560
            TabIndex        =   48
            Top             =   1155
            Width           =   2580
         End
         Begin VB.Label lblCmpSupMLCode 
            Caption         =   "CMPSUPMLCODE"
            Height          =   240
            Left            =   4560
            TabIndex        =   47
            Top             =   2235
            Width           =   2580
         End
         Begin VB.Label lblCmpCurr 
            Caption         =   "CMPCURR"
            Height          =   240
            Left            =   4560
            TabIndex        =   46
            Top             =   435
            Width           =   2580
         End
         Begin VB.Label lblCmpTiMLCode 
            Caption         =   "CMPTIMLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   1515
            Width           =   2580
         End
         Begin VB.Label lblCmpExgMLCode 
            Caption         =   "CMPEXGMLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   1155
            Width           =   2580
         End
         Begin VB.Label lblCmpRetainAC 
            Caption         =   "CMPRETAINAC"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   795
            Width           =   2580
         End
         Begin VB.Label lblCmpPayCode 
            Caption         =   "CMPPAYCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   435
            Width           =   2580
         End
         Begin VB.Label lblCmpBankAC 
            Caption         =   "CMPBANKAC"
            Height          =   255
            Left            =   -74760
            TabIndex        =   41
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblCmpBankACName 
            Caption         =   "CMPBANKACNAME"
            Height          =   240
            Left            =   -74760
            TabIndex        =   40
            Top             =   780
            Width           =   1380
         End
      End
      Begin VB.TextBox txtCmpCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   2010
      End
      Begin VB.TextBox txtCmpChiName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   7095
      End
      Begin VB.TextBox txtCmpEngName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   7080
      End
      Begin VB.Label lblCmpLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   5205
         Width           =   2460
      End
      Begin VB.Label lblCmpLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4560
         TabIndex        =   36
         Top             =   5205
         Width           =   2580
      End
      Begin VB.Label lblDspCmpLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3000
         TabIndex        =   35
         Top             =   5160
         Width           =   1425
      End
      Begin VB.Label lblDspCmpLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7200
         TabIndex        =   34
         Top             =   5160
         Width           =   1545
      End
      Begin VB.Label lblCmpCode 
         Caption         =   "CMPCODE"
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
         TabIndex        =   33
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label lblCmpChiName 
         Caption         =   "CMPCHINAME"
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   1140
         Width           =   2220
      End
      Begin VB.Label lblCmpEngName 
         Caption         =   "CMPENGNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   2220
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
            Picture         =   "CMP001.frx":3059
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":3933
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":420D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":465F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":4AB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":4DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":521D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":566F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":5989
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":5CA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":60F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CMP001.frx":69D1
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
End
Attribute VB_Name = "frmCMP001"
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

Private wiAction As Integer

Private wcCombo As Control

Private wsActNam(4) As String
Private wlKey As Long
'Row Lock Variable

Private Const wsKeyType = "MstCompany"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboCmpCode_LostFocus()
    FocusMe cboCmpCode, True
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
  
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 6480
        Me.Width = 10065
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
            Me.txtCmpEngName.Enabled = False
            Me.txtCmpChiName.Enabled = False
            Me.picCmpAdr.Enabled = False
            Me.txtCmpTel.Enabled = False
            Me.txtCmpFax.Enabled = False
            Me.txtCmpEmail.Enabled = False
            Me.txtCmpWebSite.Enabled = False
            Me.cboCmpPayCode.Enabled = False
            Me.cboCmpCurr.Enabled = False
            Me.cboCmpRetainAC.Enabled = False
            Me.cboCmpSupMLCode.Enabled = False
            Me.cboCmpExgMLCode.Enabled = False
            Me.cboCmpExlMLCode.Enabled = False
            Me.cboCmpTiMLCode.Enabled = False
            Me.cboCmpTeMLCode.Enabled = False
            Me.cboCmpDamMLCode.Enabled = False
            Me.cboCmpSamMLCode.Enabled = False
            Me.cboCmpAdjMLCode.Enabled = False
            Me.cboCmpCurrEarn.Enabled = False
            Me.txtCmpBankAC.Enabled = False
            Me.txtCmpBankACName.Enabled = False
            Me.txtCmpRemark.Enabled = False
            Me.txtCmpRptChiAdd.Enabled = False
            Me.txtCmpRptEngAdd.Enabled = False
            
            Me.cboCmpCode.Enabled = False
            Me.cboCmpCode.Visible = False
            Me.txtCmpCode.Visible = True
            Me.txtCmpCode.Enabled = False
            
            tabDetailInfo.Tab = 0
            
        Case "AfrActAdd"
            Me.cboCmpCode.Enabled = False
            Me.cboCmpCode.Visible = False
            
            Me.txtCmpCode.Enabled = True
            Me.txtCmpCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboCmpCode.Enabled = True
            Me.cboCmpCode.Visible = True
            
            Me.txtCmpCode.Enabled = False
            Me.txtCmpCode.Visible = False
            
        Case "AfrKey"
            Me.cboCmpCode.Enabled = False
            Me.txtCmpCode.Enabled = False
            
            Me.txtCmpEngName.Enabled = True
            Me.txtCmpChiName.Enabled = True
            Me.picCmpAdr.Enabled = True
            Me.txtCmpTel.Enabled = True
            Me.txtCmpFax.Enabled = True
            Me.txtCmpEmail.Enabled = True
            Me.txtCmpWebSite.Enabled = True
            Me.cboCmpPayCode.Enabled = True
            Me.cboCmpCurr.Enabled = True
            Me.cboCmpRetainAC.Enabled = True
            Me.cboCmpCurrEarn.Enabled = True
            Me.cboCmpSupMLCode.Enabled = True
            Me.cboCmpExgMLCode.Enabled = True
            Me.cboCmpExlMLCode.Enabled = True
            Me.cboCmpTiMLCode.Enabled = True
            Me.cboCmpTeMLCode.Enabled = True
            Me.cboCmpDamMLCode.Enabled = True
            Me.cboCmpSamMLCode.Enabled = True
            Me.cboCmpAdjMLCode.Enabled = True
            Me.txtCmpBankAC.Enabled = True
            Me.txtCmpBankACName.Enabled = True
            Me.txtCmpRemark.Enabled = True
            Me.txtCmpRptChiAdd.Enabled = True
            Me.txtCmpRptEngAdd.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    If Chk_txtCmpEngName = False Then
        Exit Function
    End If
    
    If Chk_txtCmpChiName = False Then
        Exit Function
    End If
    
    If Chk_txtCmpRptEngAdd = False Then
        Exit Function
    End If
    
    If Chk_txtCmpRptChiAdd = False Then
        Exit Function
    End If
    
    If Chk_txtCmpTel = False Then
        Exit Function
    End If
    
    If Chk_cboCmpCurr = False Then
        Exit Function
    End If
    
    If Chk_cboCmpRetainAC = False Then
        Exit Function
    End If
    
    If Chk_cboCmpCurrEarn = False Then
        Exit Function
    End If
    
    If Chk_cboCmpSupMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboCmpExgMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboCmpExlMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboCmpTiMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboCmpTeMLCode = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstCompany "
    wsSQL = wsSQL + "WHERE (((MstCompany.CmpCode)='" + Set_Quote(cboCmpCode.Text) + "' AND CmpStatus = '1'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "CmpID")
        Me.txtCmpEngName = ReadRs(rsRcd, "CmpEngName")
        Me.txtCmpChiName = ReadRs(rsRcd, "CmpChiName")
        Me.txtCmpAddress(1) = ReadRs(rsRcd, "CmpAddress1")
        Me.txtCmpAddress(2) = ReadRs(rsRcd, "CmpAddress2")
        Me.txtCmpAddress(3) = ReadRs(rsRcd, "CmpAddress3")
        Me.txtCmpAddress(4) = ReadRs(rsRcd, "CmpAddress4")
        Me.txtCmpTel = ReadRs(rsRcd, "CmpTel")
        Me.txtCmpFax = ReadRs(rsRcd, "CmpFax")
        Me.txtCmpEmail = ReadRs(rsRcd, "CmpEmail")
        Me.txtCmpWebSite = ReadRs(rsRcd, "CmpWebSite")
        Me.txtCmpRptChiAdd = ReadRs(rsRcd, "CmpRptChiAdd")
        Me.txtCmpRptEngAdd = ReadRs(rsRcd, "CmpRptEngAdd")
        Me.cboCmpPayCode = ReadRs(rsRcd, "CmpPayCode")
        Me.cboCmpCurr = ReadRs(rsRcd, "CmpCurr")
        Me.cboCmpRetainAC = LoadCmpRetainACCodeByID(ReadRs(rsRcd, "CmpRetainAC"))
        Me.cboCmpSupMLCode = ReadRs(rsRcd, "CmpSupMLCode")
        Me.cboCmpExgMLCode = ReadRs(rsRcd, "CmpExgMLCode")
        Me.cboCmpExlMLCode = ReadRs(rsRcd, "CmpExlMLCode")
        Me.cboCmpTiMLCode = ReadRs(rsRcd, "CmpTiMLCode")
        Me.cboCmpTeMLCode = ReadRs(rsRcd, "CmpTeMLCode")
        Me.cboCmpSamMLCode = ReadRs(rsRcd, "CmpSamMLCode")
        Me.cboCmpDamMLCode = ReadRs(rsRcd, "CmpDamMLCode")
        Me.cboCmpAdjMLCode = ReadRs(rsRcd, "CmpAdjMLCode")
        Me.cboCmpCurrEarn = ReadRs(rsRcd, "CmpCurrEarn")
        Me.txtCmpBankAC = ReadRs(rsRcd, "CmpBankAC")
        Me.txtCmpBankACName = ReadRs(rsRcd, "CmpBankACName")
        Me.txtCmpRemark = ReadRs(rsRcd, "CmpRemark")
        
        Me.lblDspCmpLastUpd = ReadRs(rsRcd, "CmpLastUpd")
        Me.lblDspCmpLastUpdDate = ReadRs(rsRcd, "CmpLastUpdDate")
        
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
    Set frmCMP001 = Nothing
End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
         If txtCmpAddress(1).Enabled Then
            txtCmpAddress(1).SetFocus
         End If
    ElseIf tabDetailInfo.Tab = 1 Then
        If txtCmpTel.Enabled Then
            txtCmpTel.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 2 Then
        If cboCmpPayCode.Enabled Then
            cboCmpPayCode.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 3 Then
        If txtCmpBankAC.Enabled Then
            txtCmpBankAC.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 4 Then
        If txtCmpRemark.Enabled Then
            txtCmpRemark.SetFocus
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
            
            Call cmdFind
            
        Case tcExit
        
            Unload Me
            
    End Select
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "CMP001"
    wsTrnCd = ""
End Sub


Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblCmpCode.Caption = Get_Caption(waScrItm, "CMPCODE")
    lblCmpEngName.Caption = Get_Caption(waScrItm, "CMPENGNAME")
    lblCmpChiName.Caption = Get_Caption(waScrItm, "CMPCHINAME")
    lblCmpAdr.Caption = Get_Caption(waScrItm, "CMPADR")
    lblCmpTel.Caption = Get_Caption(waScrItm, "CMPTEL")
    lblCmpFax.Caption = Get_Caption(waScrItm, "CMPFAX")
    lblCmpEmail.Caption = Get_Caption(waScrItm, "CMPEMAIL")
    lblCmpWebSite.Caption = Get_Caption(waScrItm, "CMPWEBSITE")
    lblCmpPayCode.Caption = Get_Caption(waScrItm, "CMPPAYCODE")
    lblCmpCurr.Caption = Get_Caption(waScrItm, "CMPCURR")
    lblCmpRetainAC.Caption = Get_Caption(waScrItm, "CMPRETAINAC")
    lblCmpSupMLCode.Caption = Get_Caption(waScrItm, "CMPSUPMLCODE")
    lblCmpExgMLCode.Caption = Get_Caption(waScrItm, "CMPEXGMLCODE")
    lblCmpExlMLCode.Caption = Get_Caption(waScrItm, "CMPEXLMLCODE")
    lblCmpTiMLCode.Caption = Get_Caption(waScrItm, "CMPTIMLCODE")
    lblCmpTeMLCode.Caption = Get_Caption(waScrItm, "CMPTEMLCODE")
    lblCmpDamMLCode.Caption = Get_Caption(waScrItm, "CMPDAMMLCODE")
    lblCmpSamMLCode.Caption = Get_Caption(waScrItm, "CMPSAMMLCODE")
    lblCmpAdjMLCode.Caption = Get_Caption(waScrItm, "CMPADJMLCODE")
    lblCmpCurrEarn.Caption = Get_Caption(waScrItm, "CMPCURREARN")
    
    lblCmpBankAC.Caption = Get_Caption(waScrItm, "CMPBANKAC")
    lblCmpBankACName.Caption = Get_Caption(waScrItm, "CMPBANKACNAME")
    lblCmpRemark.Caption = Get_Caption(waScrItm, "CMPREMARK")
    lblCmpRptChiAdd.Caption = Get_Caption(waScrItm, "CMPRPTCHIADD")
    lblCmpRptEngAdd.Caption = Get_Caption(waScrItm, "CMPRPTENGADD")
    
    lblCmpLastUpd.Caption = Get_Caption(waScrItm, "CMPLASTUPD")
    lblCmpLastUpdDate.Caption = Get_Caption(waScrItm, "CMPLASTUPDDATE")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO0")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO1")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO2")
    tabDetailInfo.TabCaption(3) = Get_Caption(waScrItm, "TABDETAILINFO3")
    tabDetailInfo.TabCaption(4) = Get_Caption(waScrItm, "TABDETAILINFO4")

    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

    wsActNam(1) = Get_Caption(waScrItm, "CMPADD")
    wsActNam(2) = Get_Caption(waScrItm, "CMPEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "CMPDELETE")
    
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
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    tblCommon.Visible = False
    Me.Caption = wsFormCaption
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtCmpCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCmpCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCmpCode.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Select Case wiAction
    
        Case CorRec, DelRec

            If LoadRecord() = False Then
                gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
                MsgBox gsMsg, vbOKOnly, gsTitle
                Exit Sub
            Else
                If RowLock(wsConnTime, wsKeyType, cboCmpCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtCmpEngName.SetFocus
End Sub

Private Function Chk_CmpCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_CmpCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT CmpStatus "
    wsSQL = wsSQL & " FROM MstCompany WHERE CmpCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "CmpStatus")
    
    Chk_CmpCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtCmpCode() As Boolean
    Dim wsStatus As String

    Chk_txtCmpCode = False
    
    If Trim(txtCmpCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpCode.SetFocus
        Exit Function
    End If

    If Chk_CmpCode(txtCmpCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "公司已存在但已無效!"
        Else
            gsMsg = "公司已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpCode.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpCode = True
End Function

Private Function Chk_cboCmpCode() As Boolean
    Dim wsStatus As String
 
    Chk_cboCmpCode = False
    
    If Trim(cboCmpCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCmpCode.SetFocus
        Exit Function
    End If

    If Chk_CmpCode(cboCmpCode.Text, wsStatus) = False Then
        gsMsg = "公司不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCmpCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "公司已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCmpCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboCmpCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmCMP001
    
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

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboCmpCode, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
   
    If wiAction = DelRec Then
        If MsgBox("你是否確定要刪除此記錄?", vbYesNo, gsTitle) = vbNo Then
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
        
    adcmdSave.CommandText = "USP_CMP001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, txtCmpCode.Text, cboCmpCode.Text))
    Call SetSPPara(adcmdSave, 4, txtCmpEngName)
    Call SetSPPara(adcmdSave, 5, txtCmpChiName)
    Call SetSPPara(adcmdSave, 6, txtCmpAddress(1))
    Call SetSPPara(adcmdSave, 7, txtCmpAddress(2))
    Call SetSPPara(adcmdSave, 8, txtCmpAddress(3))
    Call SetSPPara(adcmdSave, 9, txtCmpAddress(4))
    Call SetSPPara(adcmdSave, 10, txtCmpRptEngAdd)
    Call SetSPPara(adcmdSave, 11, txtCmpRptChiAdd)
    Call SetSPPara(adcmdSave, 12, txtCmpTel)
    Call SetSPPara(adcmdSave, 13, txtCmpFax)
    Call SetSPPara(adcmdSave, 14, txtCmpEmail)
    Call SetSPPara(adcmdSave, 15, txtCmpWebSite)
    Call SetSPPara(adcmdSave, 16, cboCmpPayCode)
    Call SetSPPara(adcmdSave, 17, cboCmpCurr)
    Call SetSPPara(adcmdSave, 18, LoadCmpRetainACIDByCode(cboCmpRetainAC))
    Call SetSPPara(adcmdSave, 19, cboCmpSupMLCode)
    Call SetSPPara(adcmdSave, 20, cboCmpExgMLCode)
    Call SetSPPara(adcmdSave, 21, cboCmpExlMLCode)
    Call SetSPPara(adcmdSave, 22, cboCmpTiMLCode)
    Call SetSPPara(adcmdSave, 23, cboCmpTeMLCode)
    Call SetSPPara(adcmdSave, 24, txtCmpBankAC)
    Call SetSPPara(adcmdSave, 25, txtCmpBankACName)
    Call SetSPPara(adcmdSave, 26, gsUserID)
    Call SetSPPara(adcmdSave, 27, wsGenDte)
    Call SetSPPara(adcmdSave, 28, txtCmpRemark)
    Call SetSPPara(adcmdSave, 29, cboCmpSamMLCode)
    Call SetSPPara(adcmdSave, 30, cboCmpDamMLCode)
    Call SetSPPara(adcmdSave, 31, cboCmpAdjMLCode)
    Call SetSPPara(adcmdSave, 32, cboCmpCurrEarn)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 33)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - CMP001!"
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

Private Sub cmdFind()
     Call OpenPromptForm
End Sub

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

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSQL As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "公司編碼"
    vFilterAry(1, 2) = "CmpCode"
    
    vFilterAry(2, 1) = "註解"
    If gsLangID <> "2" Then
        vFilterAry(2, 2) = "CmpEngName"
    Else
        vFilterAry(2, 2) = "CmpChiName"
    End If
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "編碼"
    vAry(1, 2) = "CmpCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    If gsLangID <> "2" Then
        vAry(2, 2) = "CmpEngName"
    Else
        vAry(2, 2) = "CmpChiName"
    End If
    
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        If gsLangID <> "2" Then
            wsSQL = "SELECT MstCompany.CmpCode, MstCompany.CmpEngName "
        Else
            wsSQL = "SELECT MstCompany.CmpCode, MstCompany.CmpChiName "
        End If
        wsSQL = wsSQL + "FROM MstCompany "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE MstCompany.CmpStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstCompany.CmpCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboCmpCode Then
        cboCmpCode = Trim(frmShareSearch.Tag)
        cboCmpCode.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtCmpAddress_GotFocus(Index As Integer)
    If Index = 1 Then
        If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    End If
    FocusMe txtCmpAddress(Index)
End Sub

Private Sub txtCmpAddress_KeyPress(Index As Integer, KeyAscii As Integer)
    Call chk_InpLen(txtCmpAddress(Index), 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Index < 4 Then
            txtCmpAddress(Index + 1).SetFocus
        Else
            txtCmpRptEngAdd.SetFocus
        End If
    End If
End Sub

Private Sub txtCmpAddress_LostFocus(Index As Integer)
    FocusMe txtCmpAddress(Index), True
End Sub

Private Sub txtCmpBankAC_GotFocus()
    If tabDetailInfo.Tab <> 3 Then tabDetailInfo.Tab = 3
    FocusMe txtCmpBankAC
End Sub

Private Sub txtCmpBankAC_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpBankAC, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpBankACName.SetFocus
    End If
End Sub

Private Sub txtCmpBankAC_LostFocus()
    FocusMe txtCmpBankAC, True
End Sub

Private Sub txtCmpCode_GotFocus()
    FocusMe txtCmpCode
End Sub

Private Sub txtCmpCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtCmpCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCmpCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtCmpCode_LostFocus()
    FocusMe txtCmpCode, True
End Sub

Private Sub txtCmpEngName_LostFocus()
    FocusMe txtCmpEngName, True
End Sub

Private Sub txtCmpChiName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpChiName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtCmpChiName() = True Then
            txtCmpAddress(1).SetFocus
        End If
    End If
End Sub

Private Sub txtCmpEngName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpEngName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCmpEngName() = True Then
            txtCmpChiName.SetFocus
        End If
    End If
End Sub

Private Sub txtCmpChiName_GotFocus()
    FocusMe txtCmpChiName
End Sub

Private Sub txtCmpEngName_GotFocus()
    FocusMe txtCmpEngName
End Sub

Private Function Chk_txtCmpEngName() As Boolean
    Chk_txtCmpEngName = False
    
    If Trim(txtCmpEngName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpEngName.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpEngName = True
End Function

Private Function Chk_txtCmpChiName() As Boolean
    
    Chk_txtCmpChiName = False
    
    If Trim(txtCmpChiName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpChiName.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpChiName = True
End Function

Private Sub tblCommon_DblClick()
    wcCombo.Text = tblCommon.Columns(0).Text
    
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

Private Sub cboCmpCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboCmpCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpCode
    
    If gsLangID = 1 Then
        wsSQL = "SELECT CmpCode, CmpEngName FROM MstCompany WHERE CmpStatus = '1'"
    Else
        wsSQL = "SELECT CmpCode, CmpChiName FROM MstCompany WHERE CmpStatus = '1'"
    End If
    wsSQL = wsSQL & " AND CmpCode LIKE '%" & IIf(cboCmpCode.SelLength > 0, "", Set_Quote(cboCmpCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY CmpCode "
    Call Ini_Combo(2, wsSQL, cboCmpCode.Left, cboCmpCode.Top + cboCmpCode.Height, tblCommon, "CMP001", "TBLCMP", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpCode_GotFocus()
    FocusMe cboCmpCode
End Sub

Private Sub txtCmpChiName_LostFocus()
    FocusMe txtCmpChiName, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT CmpStatus FROM MstCompany WHERE CmpCode = '" & Set_Quote(txtCmpCode) & "'"
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
        .TableKey = "CmpCode"
        .KeyLen = 10
        Set .ctlKey = txtCmpCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtCmpRemark_GotFocus()
    If tabDetailInfo.Tab <> 4 Then tabDetailInfo.Tab = 4
    FocusMe txtCmpRemark
End Sub

Private Sub txtCmpRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpRemark, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpEngName.SetFocus
    End If
End Sub

Private Sub txtCmpRemark_LostFocus()
    FocusMe txtCmpRemark, True
End Sub

Private Sub txtCmpTel_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCmpTel
End Sub

Private Sub txtCmpTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpTel, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpFax.SetFocus
    End If
End Sub

Private Sub txtCmpTel_LostFocus()
    FocusMe txtCmpTel, True
End Sub

Private Sub txtCmpFax_GotFocus()
    FocusMe txtCmpFax
End Sub

Private Sub txtCmpFax_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpEmail.SetFocus
    End If
End Sub

Private Sub txtCmpFax_LostFocus()
    FocusMe txtCmpFax, True
End Sub

Private Sub txtCmpEmail_GotFocus()
    FocusMe txtCmpEmail
End Sub

Private Sub txtCmpEmail_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpEmail, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpWebSite.SetFocus
    End If
End Sub

Private Sub txtCmpEmail_LostFocus()
    FocusMe txtCmpEmail, True
End Sub

Private Sub txtCmpWebSite_GotFocus()
    FocusMe txtCmpWebSite
End Sub

Private Sub txtCmpWebSite_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpWebSite, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboCmpPayCode.SetFocus
    End If
End Sub

Private Sub txtCmpWebSite_LostFocus()
    FocusMe txtCmpWebSite, True
End Sub

Private Sub cboCmpPayCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpPayCode
    
    wsSQL = "SELECT PayCode, PayDesc FROM MstPayTerm WHERE PayStatus = '1'"
    wsSQL = wsSQL & " AND PayCode LIKE '%" & IIf(cboCmpPayCode.SelLength > 0, "", Set_Quote(cboCmpPayCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY PayCode "
    Call Ini_Combo(2, wsSQL, cboCmpPayCode.Left, cboCmpPayCode.Top + cboCmpPayCode.Height, tblCommon, "CMP001", "TBLCMPPAY", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpPayCode_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe cboCmpPayCode
End Sub

Private Sub cboCmpPayCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpPayCode() = True Then
            cboCmpCurr.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpPayCode_LostFocus()
    FocusMe cboCmpPayCode, True
End Sub

Private Sub cboCmpCurr_DropDown()
    Dim wsSQL As String
    Dim wsCtlDte As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpCurr
    
    wsCtlDte = gsSystemDate
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboCmpCurr.SelLength > 0, "", Set_Quote(cboCmpCurr.Text)) & "%' "
    wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
    
    Call Ini_Combo(2, wsSQL, cboCmpCurr.Left, cboCmpCurr.Top + cboCmpCurr.Height, tblCommon, wsFormID, "TBLCMPCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpCurr_GotFocus()
    FocusMe cboCmpCurr
End Sub

Private Sub cboCmpCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpCurr() = True Then
            cboCmpRetainAC.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpCurr_LostFocus()
    FocusMe cboCmpCurr, True
End Sub

Private Sub cboCmpRetainAC_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpRetainAC
    
    wsSQL = "SELECT COAAccCode, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " FROM MstCOA WHERE COAStatus = '1'"
    wsSQL = wsSQL & " AND COAAccCode LIKE '%" & IIf(cboCmpRetainAC.SelLength > 0, "", Set_Quote(cboCmpRetainAC.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY COAAccCode "
    Call Ini_Combo(2, wsSQL, cboCmpRetainAC.Left, cboCmpRetainAC.Top + cboCmpRetainAC.Height, tblCommon, "CMP001", "TBLCMPRETAINAC", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpRetainAC_GotFocus()
    FocusMe cboCmpRetainAC
End Sub

Private Sub cboCmpRetainAC_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpRetainAC, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpRetainAC() = True Then
            cboCmpCurrEarn.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpRetainAC_LostFocus()
    FocusMe cboCmpRetainAC, True
End Sub

'''
Private Sub cboCmpCurrEarn_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpCurrEarn
    
    wsSQL = "SELECT COAAccCode, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " FROM MstCOA WHERE COAStatus = '1'"
    wsSQL = wsSQL & " AND COAAccCode LIKE '%" & IIf(cboCmpCurrEarn.SelLength > 0, "", Set_Quote(cboCmpCurrEarn.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY COAAccCode "
    Call Ini_Combo(2, wsSQL, cboCmpCurrEarn.Left, cboCmpCurrEarn.Top + cboCmpCurrEarn.Height, tblCommon, "CMP001", "TBLCMPCurrEarn", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpCurrEarn_GotFocus()
    FocusMe cboCmpCurrEarn
End Sub

Private Sub cboCmpCurrEarn_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpCurrEarn, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpCurrEarn() = True Then
            cboCmpExgMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpCurrEarn_LostFocus()
    FocusMe cboCmpCurrEarn, True
End Sub



Private Sub cboCmpSupMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpSupMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpSupMLCode.SelLength > 0, "", Set_Quote(cboCmpSupMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpSupMLCode.Left, cboCmpSupMLCode.Top + cboCmpSupMLCode.Height, tblCommon, "CMP001", "TBLCMPSUPML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpSupMLCode_GotFocus()
    FocusMe cboCmpSupMLCode
End Sub

Private Sub cboCmpSupMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpSupMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpSupMLCode() = True Then
            tabDetailInfo.Tab = 3
            txtCmpBankAC.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpSupMLCode_LostFocus()
    FocusMe cboCmpSupMLCode, True
End Sub

Private Sub cboCmpExgMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpExgMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpExgMLCode.SelLength > 0, "", Set_Quote(cboCmpExgMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpExgMLCode.Left, cboCmpExgMLCode.Top + cboCmpExgMLCode.Height, tblCommon, "CMP001", "TBLCMPEXGML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpExgMLCode_GotFocus()
    FocusMe cboCmpExgMLCode
End Sub

Private Sub cboCmpExgMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpExgMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpExgMLCode() = True Then
            cboCmpExlMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpExgMLCode_LostFocus()
    FocusMe cboCmpExgMLCode, True
End Sub

Private Sub cboCmpExlMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpExlMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpExlMLCode.SelLength > 0, "", Set_Quote(cboCmpExlMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpExlMLCode.Left, cboCmpExlMLCode.Top + cboCmpExlMLCode.Height, tblCommon, "CMP001", "TBLCMPEXLML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpExlMLCode_GotFocus()
    FocusMe cboCmpExlMLCode
End Sub

Private Sub cboCmpExlMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpExlMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpExlMLCode() = True Then
            cboCmpTiMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpExlMLCode_LostFocus()
    FocusMe cboCmpExlMLCode, True
End Sub

Private Sub cboCmpTiMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpTiMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpTiMLCode.SelLength > 0, "", Set_Quote(cboCmpTiMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpTiMLCode.Left, cboCmpTiMLCode.Top + cboCmpTiMLCode.Height, tblCommon, "CMP001", "TBLCMPTIML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpTiMLCode_GotFocus()
    FocusMe cboCmpTiMLCode
End Sub

Private Sub cboCmpTiMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpTiMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpTiMLCode() = True Then
            cboCmpTeMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpTiMLCode_LostFocus()
    FocusMe cboCmpTiMLCode, True
End Sub

Private Sub cboCmpTeMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpTeMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpTeMLCode.SelLength > 0, "", Set_Quote(cboCmpTeMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpTeMLCode.Left, cboCmpTeMLCode.Top + cboCmpTeMLCode.Height, tblCommon, "CMP001", "TBLCMPTEML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpTeMLCode_GotFocus()
    FocusMe cboCmpTeMLCode
End Sub

Private Sub cboCmpTeMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpTeMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpTeMLCode() = True Then
            cboCmpDamMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpTeMLCode_LostFocus()
    FocusMe cboCmpTeMLCode, True
End Sub

Private Sub txtCmpBankACName_GotFocus()
    FocusMe txtCmpBankACName
End Sub

Private Sub txtCmpBankACName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpBankACName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCmpRemark.SetFocus
    End If
End Sub

Private Sub txtCmpBankACName_LostFocus()
    FocusMe txtCmpBankACName, True
End Sub

Private Function Chk_cboCmpPayCode() As Boolean
    Chk_cboCmpPayCode = False

    If Trim(cboCmpPayCode.Text) = "" Then
        Chk_cboCmpPayCode = True
        Exit Function
    End If
    
    If Chk_CmpPayCode() = False Then
        gsMsg = "付款條款不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpPayCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpPayCode = True
End Function

Private Function Chk_CmpPayCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpPayCode = False
    
    sSQL = "SELECT MstPayTerm.PayCode FROM MstPayTerm WHERE MstPayTerm.PayCode = '" & Set_Quote(cboCmpPayCode.Text) + "' And PayStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpPayCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpCurr() As Boolean
    Chk_cboCmpCurr = False

    If Trim(cboCmpCurr.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpCurr.SetFocus
        Exit Function
    End If
    
    If Chk_Curr(cboCmpCurr, gsSystemDate) = False Then
        gsMsg = "貨幣不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpCurr.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpCurr = True
End Function

Private Function Chk_cboCmpRetainAC() As Boolean
    Chk_cboCmpRetainAC = False

    If Trim(cboCmpRetainAC.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpRetainAC.SetFocus
        Exit Function
    End If
    
    If Chk_CmpRetainAC() = False Then
        gsMsg = "盈餘帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpRetainAC.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpRetainAC = True
End Function

Private Function Chk_CmpRetainAC() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpRetainAC = False
    
    sSQL = "SELECT MstCOA.COAAccCode FROM MstCOA WHERE MstCOA.COAAccCode = '" & Set_Quote(cboCmpRetainAC.Text) + "' And COAStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpRetainAC = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpSupMLCode() As Boolean
    Chk_cboCmpSupMLCode = False

    If Trim(cboCmpSupMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpSupMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpSupMLCode() = False Then
        gsMsg = "暫記帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpSupMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpSupMLCode = True
End Function


Private Function Chk_cboCmpCurrEarn() As Boolean
    Chk_cboCmpCurrEarn = False

    If Trim(cboCmpCurrEarn.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpCurrEarn.SetFocus
        Exit Function
    End If
    
    If Chk_CmpRetainAC() = False Then
        gsMsg = "本年盈利帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpCurrEarn.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpCurrEarn = True
End Function

Private Function Chk_CmpSupMLCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpSupMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(cboCmpSupMLCode.Text) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpSupMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpExgMLCode() As Boolean
    Chk_cboCmpExgMLCode = False

    If Trim(cboCmpExgMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpExgMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpExgMLCode() = False Then
        gsMsg = "對換利益不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpExgMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpExgMLCode = True
End Function

Private Function Chk_CmpExgMLCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpExgMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(cboCmpExgMLCode.Text) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpExgMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpExlMLCode() As Boolean
    Chk_cboCmpExlMLCode = False

    If Trim(cboCmpExlMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpExlMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpExlMLCode() = False Then
        gsMsg = "對換損失不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpExlMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpExlMLCode = True
End Function

Private Function Chk_CmpExlMLCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpExlMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(cboCmpExlMLCode.Text) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpExlMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpTiMLCode() As Boolean
    Chk_cboCmpTiMLCode = False

    If Trim(cboCmpTiMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpTiMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpTiMLCode() = False Then
        gsMsg = "暫記收入不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpTiMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpTiMLCode = True
End Function

Private Function Chk_CmpTiMLCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpTiMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(cboCmpTiMLCode.Text) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CmpTiMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCmpTeMLCode() As Boolean
    Chk_cboCmpTeMLCode = False

    If Trim(cboCmpTeMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpTeMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpTeMLCode() = False Then
        gsMsg = "暫記支出不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpTeMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpTeMLCode = True
End Function

Private Function Chk_CmpTeMLCode() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpTeMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(cboCmpTeMLCode.Text) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    Chk_CmpTeMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub txtCmpRptEngAdd_GotFocus()
    FocusMe txtCmpRptEngAdd
End Sub

Private Sub txtCmpRptEngAdd_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpRptEngAdd, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtCmpRptEngAdd() = True Then
            txtCmpRptChiAdd.SetFocus
        End If
    End If
End Sub

Private Sub txtCmpRptChiAdd_LostFocus()
    FocusMe txtCmpRptChiAdd, True
End Sub

Private Sub txtCmpRptChiAdd_GotFocus()
    FocusMe txtCmpRptChiAdd
End Sub

Private Sub txtCmpRptChiAdd_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCmpRptChiAdd, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtCmpRptChiAdd() = True Then
            txtCmpTel.SetFocus
        End If
    End If
End Sub

Private Function Chk_txtCmpRptChiAdd() As Boolean
    Chk_txtCmpRptChiAdd = False
    
    If Trim(txtCmpRptChiAdd.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpRptChiAdd.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpRptChiAdd = True
End Function

Private Function Chk_txtCmpRptEngAdd() As Boolean
    Chk_txtCmpRptEngAdd = False
    
    If Trim(txtCmpRptEngAdd.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpRptEngAdd.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpRptEngAdd = True
End Function

Private Function Chk_txtCmpTel() As Boolean
    Chk_txtCmpTel = False
    
    If Trim(txtCmpTel.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCmpTel.SetFocus
        Exit Function
    End If
    
    Chk_txtCmpTel = True
End Function

Public Function LoadCmpRetainACCodeByID(ByVal inCode) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT COAAccCode FROM MstCOA WHERE COAAccID =" & To_Value(inCode)
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
        LoadCmpRetainACCodeByID = ReadRs(rsRcd, "COAAccCode")
    Else
        LoadCmpRetainACCodeByID = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function LoadCmpRetainACIDByCode(ByVal inCode) As Long
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT COAAccID FROM MstCOA WHERE COAAccCode='" & Set_Quote(inCode) & "' AND COAStatus='1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
        LoadCmpRetainACIDByCode = ReadRs(rsRcd, "COAAccID")
    Else
        LoadCmpRetainACIDByCode = 0
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub cboCmpDamMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpDamMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpDamMLCode.SelLength > 0, "", Set_Quote(cboCmpDamMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpDamMLCode.Left + tabDetailInfo.Left, cboCmpDamMLCode.Top + tabDetailInfo.Top + cboCmpDamMLCode.Height, tblCommon, "CMP001", "TBLCMPDAMML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpDamMLCode_GotFocus()
    tabDetailInfo.Tab = 2
    FocusMe cboCmpDamMLCode
End Sub

Private Sub cboCmpDamMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpDamMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpDamMLCode() = True Then
            cboCmpSamMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpDamMLCode_LostFocus()
    FocusMe cboCmpDamMLCode, True
End Sub

Private Sub cboCmpSamMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpSamMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpSamMLCode.SelLength > 0, "", Set_Quote(cboCmpSamMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpSamMLCode.Left + tabDetailInfo.Left, cboCmpSamMLCode.Top + tabDetailInfo.Top + cboCmpSamMLCode.Height, tblCommon, "CMP001", "TBLCMPDAMML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpSamMLCode_GotFocus()
    tabDetailInfo.Tab = 2
    FocusMe cboCmpSamMLCode
End Sub

Private Sub cboCmpSamMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpSamMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpSamMLCode() = True Then
            cboCmpAdjMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpSamMLCode_LostFocus()
    FocusMe cboCmpSamMLCode, True
End Sub

Private Sub cboCmpAdjMLCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCmpAdjMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboCmpAdjMLCode.SelLength > 0, "", Set_Quote(cboCmpAdjMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCmpAdjMLCode.Left + tabDetailInfo.Left, cboCmpAdjMLCode.Top + tabDetailInfo.Top + cboCmpAdjMLCode.Height, tblCommon, "CMP001", "TBLCMPDAMML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCmpAdjMLCode_GotFocus()
    tabDetailInfo.Tab = 2
    FocusMe cboCmpAdjMLCode
End Sub

Private Sub cboCmpAdjMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCmpAdjMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCmpTeMLCode() = True Then
            cboCmpSupMLCode.SetFocus
        End If
    End If
End Sub

Private Sub cboCmpAdjMLCode_LostFocus()
    FocusMe cboCmpAdjMLCode, True
End Sub

Private Function Chk_cboCmpDamMLCode() As Boolean
    Chk_cboCmpDamMLCode = False

    If Trim(cboCmpDamMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpDamMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpMLCode(cboCmpDamMLCode) = False Then
        gsMsg = "報損帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpDamMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpDamMLCode = True
End Function

Private Function Chk_cboCmpSamMLCode() As Boolean
    Chk_cboCmpSamMLCode = False

    If Trim(cboCmpSamMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpSamMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpMLCode(cboCmpSamMLCode) = False Then
        gsMsg = "樣本帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpSamMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpSamMLCode = True
End Function

Private Function Chk_cboCmpAdjMLCode() As Boolean
    Chk_cboCmpAdjMLCode = False

    If Trim(cboCmpAdjMLCode.Text) = "" Then
        gsMsg = "沒有輸入資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpAdjMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_CmpMLCode(cboCmpAdjMLCode) = False Then
        gsMsg = "調整帳戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCmpAdjMLCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCmpAdjMLCode = True
End Function

Private Function Chk_CmpMLCode(sMLCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_CmpMLCode = False
    
    sSQL = "SELECT MstMerchClass.MLCode FROM MstMerchClass WHERE MstMerchClass.MLCode = '" & Set_Quote(sMLCode) + "' And MLStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CmpMLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

