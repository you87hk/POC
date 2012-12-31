VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAP003 
   Caption         =   "訂貨單"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "frmAP003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "frmAP003.frx":030A
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtChqAmtOrg 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmAP003.frx":2A0D
      Top             =   2520
      Width           =   1900
   End
   Begin VB.ComboBox cboTMPML 
      Height          =   300
      Left            =   1400
      TabIndex        =   9
      Top             =   3480
      Width           =   1900
   End
   Begin VB.Frame fraRemain 
      Height          =   975
      Left            =   0
      TabIndex        =   37
      Top             =   2880
      Width           =   9735
      Begin VB.TextBox txtRemAmt 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1410
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmAP003.frx":2A1C
         Top             =   240
         Width           =   1900
      End
      Begin VB.Label lblDspTmpMlDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3360
         TabIndex        =   44
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblTmpMl 
         Caption         =   "MLCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LblDspRemDte 
         BorderStyle     =   1  '單線固定
         Caption         =   "mm/dd/yyyy"
         Height          =   285
         Left            =   8460
         TabIndex        =   42
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblChqDte 
         Caption         =   "Cheque Date"
         Height          =   180
         Left            =   7290
         TabIndex        =   41
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblDspRemAmtl 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9, 999,999,999.99"
         Height          =   255
         Left            =   5205
         TabIndex        =   40
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label lblRemAmtLoc 
         Caption         =   "Cheque Local Amount"
         Height          =   195
         Left            =   3645
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblRemAmtOrg 
         Caption         =   "Outstanding Cheque Amount"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtRmk 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Text            =   "012345678901234578901234567890123457890123456789"
      Top             =   2160
      Width           =   5535
   End
   Begin VB.ComboBox cboMLCode 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   2010
   End
   Begin VB.ComboBox cboCurr 
      Height          =   300
      Left            =   8280
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtExcr 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   8280
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox cboVdrCode 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin MSMask.MaskEdBox medDocDate 
      Height          =   285
      Left            =   8280
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   5760
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
            Picture         =   "frmAP003.frx":2A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":330A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":3BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":4036
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":4488
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":47A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":4BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":5046
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":5360
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":567A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":5ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP003.frx":63A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   17
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
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   2655
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4683
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Header Information"
      TabPicture(0)   =   "frmAP003.frx":66D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDspInvAmt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDspUnInvAmt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUnInvAmt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInvAmt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tblInvoice"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmAP003.frx":66EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblUnOthAmt"
      Tab(1).Control(1)=   "lblDspUnOthAmt"
      Tab(1).Control(2)=   "lblDspOthAmt"
      Tab(1).Control(3)=   "lblOthAmt"
      Tab(1).Control(4)=   "tblDetail"
      Tab(1).ControlCount=   5
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   1815
         Left            =   -74880
         OleObjectBlob   =   "frmAP003.frx":6708
         TabIndex        =   11
         Top             =   120
         Width           =   9495
      End
      Begin TrueDBGrid60.TDBGrid tblInvoice 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmAP003.frx":C1C3
         TabIndex        =   10
         Top             =   120
         Width           =   9495
      End
      Begin VB.Label lblInvAmt 
         Caption         =   "NETAMTORG"
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label lblUnInvAmt 
         Caption         =   "NETAMTLOC"
         Height          =   255
         Left            =   6480
         TabIndex        =   35
         Top             =   1980
         Width           =   1755
      End
      Begin VB.Label lblDspUnInvAmt 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   8280
         TabIndex        =   34
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label lblDspInvAmt 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5040
         TabIndex        =   33
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label lblOthAmt 
         Caption         =   "NETAMTORG"
         Height          =   255
         Left            =   -71880
         TabIndex        =   32
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label lblDspOthAmt 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -69960
         TabIndex        =   28
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label lblDspUnOthAmt 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -66720
         TabIndex        =   27
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label lblUnOthAmt 
         Caption         =   "NETAMTLOC"
         Height          =   255
         Left            =   -68520
         TabIndex        =   26
         Top             =   1980
         Width           =   1755
      End
   End
   Begin VB.Label lblChqAmtOrg 
      Caption         =   "Cheque Amount (Org)"
      Height          =   195
      Left            =   120
      TabIndex        =   47
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label lblChqAmtLoc 
      Caption         =   "Cheque Amount (Org)"
      Height          =   195
      Left            =   3645
      TabIndex        =   46
      Top             =   2550
      Width           =   1575
   End
   Begin VB.Label LblDspChqAmtLoc 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "999,999,999.99"
      Height          =   255
      Left            =   5205
      TabIndex        =   45
      Top             =   2520
      Width           =   1770
   End
   Begin VB.Label lblRmk 
      Caption         =   "RMK"
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblDspMLDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   3480
      TabIndex        =   30
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lblMlCode 
      Caption         =   "MLCODE"
      Height          =   240
      Left            =   120
      TabIndex        =   29
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lblVdrTel 
      Caption         =   "CUSTEL"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label lblVdrFax 
      Caption         =   "CUSFAX"
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label lblDspVdrFax 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4920
      TabIndex        =   22
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblVdrName 
      Caption         =   "CUSNAME"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label lblDspVdrTel 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1440
      TabIndex        =   20
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblExcr 
      Caption         =   "EXCR"
      Height          =   255
      Left            =   7005
      TabIndex        =   19
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Label LblCurr 
      Caption         =   "CURR"
      Height          =   255
      Left            =   7005
      TabIndex        =   18
      Top             =   1140
      Width           =   1200
   End
   Begin VB.Label lblDspVdrName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblVdrCode 
      Caption         =   "CUSCODE"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label lblDocDate 
      Caption         =   "DOCDATE"
      Height          =   255
      Left            =   7005
      TabIndex        =   13
      Top             =   780
      Width           =   1200
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
      TabIndex        =   12
      Top             =   420
      Width           =   1215
   End
   Begin VB.Menu mnuIPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuIPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuOPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuOPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAP003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private waResult As New XArrayDB
Private waInvoice As New XArrayDB
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private waPopUpSub As New XArrayDB
Private wcCombo As Control




Private wsOldVdrNo As String
Private wsOldCurCd As String
Private wsTmpMl As String
Private wbLocked As Boolean
Private wsCtlDte As String

Private wbReadOnly As Boolean

Private Const IINVNO = 0
Private Const ILINE = 1
Private Const ICURR = 2
Private Const IOSAMT = 3
Private Const ISETAMTORG = 4
Private Const ISETAMTLOC = 5
Private Const IEXCR = 6
Private Const IIPDTID = 7
Private Const IDummy = 8




Private Const OMLCODE = 0
Private Const OCURR = 1
Private Const OEXCR = 2
Private Const OOTHAMTORG = 3
Private Const OOTHAMTLOC = 4
Private Const ODummy = 5



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
Private wlVdrID As Long
Private wlSaleID As Long


Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "APSTHD"
Private Const wsInvKeyType = "APIPHD"   'Use for the locking

Private wsFormID As String
Private wsUsrId As String

Private wsSrcCd As String

Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String
Private wsBaseExcr As String





Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waInvoice.ReDim 0, -1, IINVNO, IIPDTID
    Set tblInvoice.Array = waInvoice
    tblInvoice.ReBind
    tblInvoice.Bookmark = 0
    
    waResult.ReDim 0, -1, OMLCODE, OOTHAMTLOC
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
    Call SetFieldStatus("RemFalse")
    
    
    Call SetDateMask(medDocDate)
    
    wsCtlDte = getCtrlMth("AP")
      
    
    wsOldVdrNo = ""
    wsOldCurCd = ""
    wsTmpMl = ""
    
    wlKey = 0
    wlVdrID = 0
    wbReadOnly = False
    
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
            cboMLCode.SetFocus
           End If
        End If
    End If
    
End Sub

Private Sub cboCurr_DropDown()
    
    Dim wsSQL As String
  
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCurr
    
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboCurr.SelLength > 0, "", Set_Quote(cboCurr.Text)) & "%' "
    wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Right(wsCtlDte, 2)) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Left(wsCtlDte, 4)) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
    Call Ini_Combo(2, wsSQL, cboCurr.Left, cboCurr.Top + cboCurr.Height, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
    
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
  
    
    wsSQL = "SELECT APSHDOCNO, APSHDOCDATE "
    wsSQL = wsSQL & " FROM APSTHD "
    wsSQL = wsSQL & " WHERE APSHDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
'    wsSql = wsSql & " AND APSHUPDFLG = 'N'"
'    wsSql = wsSql & " AND APSHSTATUS = '1'"
'    wsSql = wsSql & " AND APSHPGMNO = '" & wsFormID & "' "
    wsSQL = wsSQL & " ORDER BY APSHDOCNO, APSHDOCDATE "
  
    
    Call Ini_Combo(2, wsSQL, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
Dim rsRcd As New ADODB.Recordset
Dim wsSQL As String
Dim wsStatus As String
Dim wsUpdFlg As String
Dim wsTrnCode As String
Dim wsPgmNo As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoGen("PV") = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNo.SetFocus
        Exit Function
    End If
    
        
   If Chk_DocNo(cboDocNo, wsStatus, wsUpdFlg, wsPgmNo) = True Then
        
        If wsPgmNo <> wsFormID Then
            gsMsg = "This is not a valid form code!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
        End If
        
        If wsStatus = "4" Or wsUpdFlg = "Y" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
            Chk_cboDocNo = True
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
        End If
    

    
    Else
    
        wsSQL = "SELECT APCQCHQID FROM APCHEQUE "
        wsSQL = wsSQL & " WHERE APCQCHQNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND APCQPGMNO <> '" & wsFormID & "' "
    
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used by Cheque!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
        rsRcd.Close
        Set rsRcd = Nothing
   
        wsSQL = "SELECT IPHDDOCNO FROM APIPHD "
        wsSQL = wsSQL & " WHERE IPHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND IPHDPGMNO <> '" & wsFormID & "' "
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used by Invoice!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
        rsRcd.Close
        Set rsRcd = Nothing
    
        
    End If
    
    
    Chk_cboDocNo = True

End Function




Private Sub Ini_Scr_AfrKey()
    
    
    
    If LoadRecord() = False Then
        wiAction = AddRec
        medDocDate.Text = Dsp_Date(Now)
        Call SetButtonStatus("AfrKeyAdd")
        Call SetFieldStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        wbLocked = False
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
            wbLocked = True
        End If
        wsOldVdrNo = cboVdrCode.Text
        wsOldCurCd = cboCurr.Text
        
    
        If UCase(cboCurr) = UCase(wsBaseCurCd) Then
            txtExcr.Text = Format("1", gsExrFmt)
            txtExcr.Enabled = False
        Else
            txtExcr.Enabled = True
        End If
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKeyEdit")
        
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    
    If cboVdrCode.Enabled = False Then
    cboCurr.SetFocus
    Else
    cboVdrCode.SetFocus
    End If
        
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
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
   
    wsSQL = "SELECT APSHDOCID, APSHDOCNO, APSHVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, "
    wsSQL = wsSQL & "APSHDOCDATE, APSHCURR, APSHEXCR, APSHCATCODE, APSHMLCODE, "
    wsSQL = wsSQL & "APSHCHQAMT, APSHCHQAMTL, APSHREMARK "
    wsSQL = wsSQL & "FROM  APSTHD, mstVENDOR "
    wsSQL = wsSQL & "WHERE APSHDOCNO = '" & Set_Quote(cboDocNo) & "' "
    wsSQL = wsSQL & "AND APSHVDRID = VDRID "
    wsSQL = wsSQL & "AND APSHPGMNO = '" & wsFormID & "'"
    wsSQL = wsSQL & "AND APSHSTATUS <> '2'"
    
        
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    wlKey = ReadRs(rsRcd, "APSHDOCID")
    medDocDate.Text = ReadRs(rsRcd, "APSHDOCDATE")
    wlVdrID = ReadRs(rsRcd, "VDRID")
    cboVdrCode.Text = ReadRs(rsRcd, "VDRCODE")
    lblDspVdrName.Caption = ReadRs(rsRcd, "VDRNAME")
    lblDspVdrTel.Caption = ReadRs(rsRcd, "VDRTEL")
    lblDspVdrFax.Caption = ReadRs(rsRcd, "VDRFAX")
    cboCurr.Text = ReadRs(rsRcd, "APSHCURR")
    txtExcr.Text = Format(ReadRs(rsRcd, "APSHEXCR"), gsExrFmt)
    txtChqAmtOrg.Text = NBRnd(ReadRs(rsRcd, "APSHCHQAMT"), giAmtDp)
    If To_Value(txtChqAmtOrg.Text) > 0 Then Call SetFieldStatus("RemTrue")
    LblDspChqAmtLoc.Caption = NBRnd(ReadRs(rsRcd, "APSHCHQAMTL"), giAmtDp)
    
   
    cboMLCode = ReadRs(rsRcd, "APSHMLCODE")
    txtRmk = ReadRs(rsRcd, "APSHREMARK")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    rsRcd.Close
    
    
    wsSQL = " SELECT IPDTID, IPHDDOCNO, IPDTDOCLINE, APSDCURR, IPDTBALAMT AS BAL, "
    wsSQL = wsSQL & " APSDTRNAMT * -1 AS TRNAMT, APSDTRNAMTL * -1 AS TRNAMTL, IPHDEXCR "
    wsSQL = wsSQL & " FROM  APSTDT, APIPDT, APIPHD "
    wsSQL = wsSQL & " WHERE APSDDOCID  = " & wlKey
    wsSQL = wsSQL & " AND   APSDLNTYP  = 'I' "
    wsSQL = wsSQL & " AND   APSDIPDTID = IPDTID "
    wsSQL = wsSQL & " AND   IPHDDOCID = IPDTDOCID "
    wsSQL = wsSQL & " ORDER BY APSDDOCLINE "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
    
    rsRcd.MoveFirst
    With waInvoice
         .ReDim 0, -1, IINVNO, IIPDTID
         Do While Not rsRcd.EOF
             .AppendRows
             waInvoice(.UpperBound(1), IINVNO) = ReadRs(rsRcd, "IPHDDOCNO")
             waInvoice(.UpperBound(1), ILINE) = Format(To_Value(ReadRs(rsRcd, "IPDTDOCLINE")), "000")
             waInvoice(.UpperBound(1), ICURR) = ReadRs(rsRcd, "APSDCURR")
             waInvoice(.UpperBound(1), IEXCR) = ReadRs(rsRcd, "IPHDEXCR")
             waInvoice(.UpperBound(1), IOSAMT) = NBRnd(To_Value(ReadRs(rsRcd, "BAL")) + To_Value(ReadRs(rsRcd, "TRNAMT")), giAmtDp)
             waInvoice(.UpperBound(1), ISETAMTORG) = Format(ReadRs(rsRcd, "TRNAMT"), gsAmtFmt)
             waInvoice(.UpperBound(1), ISETAMTLOC) = Format(ReadRs(rsRcd, "TRNAMTL"), gsAmtFmt)
             waInvoice(.UpperBound(1), IIPDTID) = ReadRs(rsRcd, "IPDTID")
             
            If NoMoreInvNo(waInvoice, ReadRs(rsRcd, "IPHDDOCNO"), .UpperBound(1)) Then
                  If RowLock(wsConnTime, wsInvKeyType, ReadRs(rsRcd, "IPHDDOCNO"), wsFormID, wsUsrId) = False Then
                      gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                      MsgBox gsMsg, vbOKOnly, gsTitle
                  End If
            End If
                
             rsRcd.MoveNext
         Loop
    End With
    tblInvoice.ReBind
    tblInvoice.FirstRow = 0
    
    End If
    rsRcd.Close
    
    
    wsSQL = " SELECT APSDMLCODE, APSDDOCLINE, APSDCURR, APSDEXCR, "
    wsSQL = wsSQL & " APSDTRNAMT, APSDTRNAMTL "
    wsSQL = wsSQL & " FROM  APSTDT "
    wsSQL = wsSQL & " WHERE APSDDOCID  = " & wlKey
    wsSQL = wsSQL & " AND   APSDLNTYP  = 'O' "
    wsSQL = wsSQL & " ORDER BY APSDDOCLINE "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, OMLCODE, OOTHAMTLOC
         Do While Not rsRcd.EOF
             .AppendRows
             waResult(.UpperBound(1), OMLCODE) = ReadRs(rsRcd, "APSDMLCODE")
             waResult(.UpperBound(1), OCURR) = ReadRs(rsRcd, "APSDCURR")
             waResult(.UpperBound(1), OEXCR) = ReadRs(rsRcd, "APSDEXCR")
             waResult(.UpperBound(1), OOTHAMTORG) = Format(To_Value(ReadRs(rsRcd, "APSDTRNAMT") * -1), gsAmtFmt)
             waResult(.UpperBound(1), OOTHAMTLOC) = Format(To_Value(ReadRs(rsRcd, "APSDTRNAMTL")) * -1, gsAmtFmt)
             
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    
    End If
    
    rsRcd.Close
    
    
    wsSQL = " SELECT APSDMLCODE, IPHDMLCODE, "
    wsSQL = wsSQL & " APSDTRNAMT, APSDTRNAMTL "
    wsSQL = wsSQL & " FROM  APSTHD, APSTDT, APIPHD "
    wsSQL = wsSQL & " WHERE APSDDOCID  = " & wlKey
    wsSQL = wsSQL & " AND APSDLNTYP  = 'R' "
    wsSQL = wsSQL & " AND APSHDOCID  = APSDDOCID "
    wsSQL = wsSQL & " AND APSHDOCNO  = IPHDDOCNO "
        
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
       txtRemAmt.Text = NBRnd(0 - NBRnd(ReadRs(rsRcd, "APSDTRNAMT"), giAmtDp), giAmtDp)
       lblDspRemAmtl.Caption = NBRnd(0 - NBRnd(ReadRs(rsRcd, "APSDTRNAMTL"), giAmtDp), giAmtDp)
       cboTMPML.Text = ReadRs(rsRcd, "IPHDMLCODE")
       LblDspRemDte = medDocDate.Text
       
    Else
       txtRemAmt.Text = NBRnd("0", giAmtDp)
       cboTMPML.Text = ""
      
    End If
    
    lblDspTmpMlDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboTMPML.Text) & "'", "MLDESC")
    
    
    rsRcd.Close
    Set rsRcd = Nothing
    
     Call Calc_ChargeTotal
    Call Calc_InvTotal
   
    
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblVdrName.Caption = Get_Caption(waScrItm, "VDRNAME")
    lblVdrTel.Caption = Get_Caption(waScrItm, "VDRTEL")
    lblVdrFax.Caption = Get_Caption(waScrItm, "VDRFAX")
    lblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblTmpMl.Caption = Get_Caption(waScrItm, "TMPML")
    
    lblChqAmtOrg.Caption = Get_Caption(waScrItm, "CHQAMTORG")
    lblChqAmtLoc.Caption = Get_Caption(waScrItm, "CHQAMTLOC")
    lblRemAmtOrg.Caption = Get_Caption(waScrItm, "REMAMTORG")
    lblRemAmtLoc.Caption = Get_Caption(waScrItm, "REMAMTLOC")
    lblChqDte.Caption = Get_Caption(waScrItm, "CHQDATE")
    
    lblInvAmt.Caption = Get_Caption(waScrItm, "INVAMT")
    lblUnInvAmt.Caption = Get_Caption(waScrItm, "UNINVAMT")
    lblOthAmt.Caption = Get_Caption(waScrItm, "OTHAMT")
    lblUnOthAmt.Caption = Get_Caption(waScrItm, "UNOTHAMT")
    
    With tblInvoice
        .Columns(IINVNO).Caption = Get_Caption(waScrItm, "IINVNO")
        .Columns(ILINE).Caption = Get_Caption(waScrItm, "ILINE")
        .Columns(ICURR).Caption = Get_Caption(waScrItm, "ICURR")
        .Columns(IOSAMT).Caption = Get_Caption(waScrItm, "IOSAMT")
        .Columns(ISETAMTORG).Caption = Get_Caption(waScrItm, "ISETAMTORG")
        .Columns(ISETAMTLOC).Caption = Get_Caption(waScrItm, "ISETAMTLOC")
    End With
    
    
    With tblDetail
        .Columns(OMLCODE).Caption = Get_Caption(waScrItm, "OMLCODE")
        .Columns(OCURR).Caption = Get_Caption(waScrItm, "OCURR")
        .Columns(OEXCR).Caption = Get_Caption(waScrItm, "OEXCR")
        .Columns(OOTHAMTORG).Caption = Get_Caption(waScrItm, "OOTHAMTORG")
        .Columns(OOTHAMTLOC).Caption = Get_Caption(waScrItm, "OOTHAMTLOC")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    
    lblRmk.Caption = Get_Caption(waScrItm, "RMK")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "ARADD")
    wsActNam(2) = Get_Caption(waScrItm, "AREDIT")
    wsActNam(3) = Get_Caption(waScrItm, "ARELETE")
    
    Call Ini_PopMenu(mnuIPopUpSub, "POPUP", waPopUpSub)
    Call Ini_PopMenu(mnuOPopUpSub, "POPUP", waPopUpSub)
    
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
    Set waInvoice = Nothing
    Set waScrToolTip = Nothing
    Set waScrItm = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmAP003 = Nothing

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
    
    
    If Chk_ValidDocDate(medDocDate.Text, "AP") = False Then
        medDocDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medDocDate = True

End Function








Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
       If tblInvoice.Enabled Then
            tblInvoice.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
    
        If tblDetail.Enabled Then
            tblDetail.SetFocus
        End If
        
   
    End If
End Sub



Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
         wcCombo.Text = tblCommon.Columns(0).Text
       
    ElseIf wcCombo.Name = tblInvoice.Name Then
            tblInvoice.EditActive = True
            wcCombo.Text = tblCommon.Columns(0).Text
             
            If tblInvoice.Col = IINVNO Then
                tblInvoice.Columns(ILINE).Text = tblCommon.Columns(1).Text
            End If
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
       
        ElseIf wcCombo.Name = tblInvoice.Name Then
            tblInvoice.EditActive = True
            wcCombo.Text = tblCommon.Columns(0).Text
             
            If tblInvoice.Col = IINVNO Then
                tblInvoice.Columns(ILINE).Text = tblCommon.Columns(1).Text
            End If
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
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT APSHSTATUS FROM APSTHD WHERE APSHDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

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
    tblInvoice.Enabled = True
    
    Chk_KeyFld = True

End Function
Private Function cmdSave() As Boolean

    Dim wsGenDte As String
    Dim wsDteTim As String
    
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlRowCtr As Long
    Dim wsCtlPrd As String
    Dim wsSts As String
    Dim i As Integer
    Dim wdTmpAmt As Double
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    wsDteTim = Change_SQLDate(Now)
    
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbReadOnly Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
   
    'check each detail line
    If Chk_LockedRecords Then
        MousePointer = vbDefault
        Exit Function
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
    
 
    
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_AP003A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 4, wlVdrID)
    Call SetSPPara(adcmdSave, 5, medDocDate.Text)
    Call SetSPPara(adcmdSave, 6, cboCurr.Text)
    Call SetSPPara(adcmdSave, 7, txtExcr.Text)
    Call SetSPPara(adcmdSave, 8, To_Value(txtChqAmtOrg.Text))
    Call SetSPPara(adcmdSave, 9, NBRnd(To_Value(txtExcr.Text) * To_Value(txtChqAmtOrg.Text), giAmtDp))
    Call SetSPPara(adcmdSave, 10, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 11, "")
    Call SetSPPara(adcmdSave, 12, txtRmk.Text)
    Call SetSPPara(adcmdSave, 13, wsFormID)
    Call SetSPPara(adcmdSave, 14, gsUserID)
    Call SetSPPara(adcmdSave, 15, wsGenDte)
    Call SetSPPara(adcmdSave, 16, wsDteTim)
    
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 17)
    wsDocNo = GetSPPara(adcmdSave, 18)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waInvoice.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_AP003B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waInvoice.UpperBound(1)
            If IsEmptyInvRow(wiCtr) = False Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waInvoice(wiCtr, IIPDTID))
                Call SetSPPara(adcmdSave, 4, wiCtr + 1)
                Call SetSPPara(adcmdSave, 5, cboMLCode.Text)
                Call SetSPPara(adcmdSave, 6, "")
                Call SetSPPara(adcmdSave, 7, "")
                Call SetSPPara(adcmdSave, 8, "")
                Call SetSPPara(adcmdSave, 9, waInvoice(wiCtr, ICURR))
                Call SetSPPara(adcmdSave, 10, "")
                Call SetSPPara(adcmdSave, 11, waInvoice(wiCtr, ISETAMTORG))
                Call SetSPPara(adcmdSave, 12, waInvoice(wiCtr, ISETAMTLOC))
                Call SetSPPara(adcmdSave, 13, "")
                Call SetSPPara(adcmdSave, 14, medDocDate.Text)
                Call SetSPPara(adcmdSave, 15, waInvoice(wiCtr, IEXCR))
                Call SetSPPara(adcmdSave, 16, "I")
                Call SetSPPara(adcmdSave, 17, wsFormID)
                Call SetSPPara(adcmdSave, 18, gsUserID)
                Call SetSPPara(adcmdSave, 19, wsGenDte)
                
                adcmdSave.Execute
            End If
        Next
    End If
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_AP003B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If IsEmptyRow(wiCtr) = False Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, "")
                Call SetSPPara(adcmdSave, 4, wiCtr + 1)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, OMLCODE))
                Call SetSPPara(adcmdSave, 6, "")
                Call SetSPPara(adcmdSave, 7, "")
                Call SetSPPara(adcmdSave, 8, "")
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, OCURR))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, OEXCR))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, OOTHAMTORG))
                Call SetSPPara(adcmdSave, 12, NBRnd(To_Value(waResult(wiCtr, OOTHAMTORG) * To_Value(waResult(wiCtr, OEXCR))), giAmtDp))
                Call SetSPPara(adcmdSave, 13, "")
                Call SetSPPara(adcmdSave, 14, medDocDate.Text)
                Call SetSPPara(adcmdSave, 15, "")
                Call SetSPPara(adcmdSave, 16, "O")
                Call SetSPPara(adcmdSave, 17, wsFormID)
                Call SetSPPara(adcmdSave, 18, gsUserID)
                Call SetSPPara(adcmdSave, 19, wsGenDte)
                
                adcmdSave.Execute
            End If
        Next
    End If
    
    
    
    If To_Value(txtRemAmt.Text) <> 0 Then
    
        adcmdSave.CommandText = "USP_AP003B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, "")
                Call SetSPPara(adcmdSave, 4, 1)
                Call SetSPPara(adcmdSave, 5, cboTMPML.Text)
                Call SetSPPara(adcmdSave, 6, "")
                Call SetSPPara(adcmdSave, 7, "")
                Call SetSPPara(adcmdSave, 8, "")
                Call SetSPPara(adcmdSave, 9, cboCurr.Text)
                Call SetSPPara(adcmdSave, 10, txtExcr.Text)
                Call SetSPPara(adcmdSave, 11, txtRemAmt.Text)
                Call SetSPPara(adcmdSave, 12, lblDspRemAmtl)
                Call SetSPPara(adcmdSave, 13, "")
                Call SetSPPara(adcmdSave, 14, medDocDate.Text)
                Call SetSPPara(adcmdSave, 15, "")
                Call SetSPPara(adcmdSave, 16, "R")
                Call SetSPPara(adcmdSave, 17, wsFormID)
                Call SetSPPara(adcmdSave, 18, gsUserID)
                Call SetSPPara(adcmdSave, 19, wsGenDte)
                
        adcmdSave.Execute
        
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
    
    
    If To_Value(lblDspUnInvAmt.Caption) <> 0 Then
        gsMsg = "Undistributed Amount must equal 0!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboVdrCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboMLCode Then Exit Function
    If Not Chk_txtChqAmtOrg(txtChqAmtOrg.Text) Then Exit Function
    If Not Chk_txtRemAmt(txtRemAmt.Text) Then Exit Function
    If To_Value(txtRemAmt.Text) <> 0 Then
       If Not Chk_cboTmpML Then Exit Function
    End If
    
    tabDetailInfo.Tab = 0
    
    If waInvoice.UpperBound(1) < 0 Then
        gsMsg = "No Invoice Information!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       tblInvoice.SetFocus
       Exit Function
    End If
    
    With tblInvoice
        If .EditActive = True Then Exit Function
        .Update
        If Chk_InvGrdRow(.FirstRow + .Row) = False Then
            .SetFocus
            Exit Function
        End If
    End With
    
    
    
    If Chk_NoDup(To_Value(tblInvoice.Bookmark)) = False Then
        tblInvoice.FirstRow = tblInvoice.Row
        tblInvoice.Col = IINVNO
        tblInvoice.SetFocus
        Exit Function
    End If
    
    With tblDetail
        If .EditActive = True Then Exit Function
        .Update
        If Chk_GrdRow(To_Value(.FirstRow) + .Row) = False Then
            tabDetailInfo.Tab = 1
            .SetFocus
            Exit Function
        End If
    End With
    
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
    


Private Sub cmdNew()

    Dim newForm As New frmAP003
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmAP003
    
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
    wsFormID = "AP003"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    If getExcRate(wsBaseCurCd, wsConnTime, wsBaseExcr, "") = False Then
            wsBaseExcr = Format(1, gsExrFmt)
    End If
    wsSrcCd = "AP"

    
    


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


Private Sub tblDetail_OnAddNew()
    
    With tblDetail
    If Trim(.Columns(OCURR)) = "" Then
        .Columns(OCURR) = wsBaseCurCd
        .Columns(OEXCR) = wsBaseExcr
    End If
    End With
    
    Call chk_BaseCurr
End Sub

Private Sub tblInvoice_AfterColUpdate(ByVal ColIndex As Integer)
   
    On Error GoTo tblInvoice_AfterColUpdate_Err
        
    With tblInvoice
        .Update
    End With
    
    Exit Sub
    
tblInvoice_AfterColUpdate_Err:
    MsgBox "tblInvoice_AfterColUpdate_Err"
    
End Sub
Private Sub tblInvoice_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    
    Dim wsExcr  As String
    
    On Error GoTo Tbl_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblInvoice.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
    
    With tblInvoice
        Select Case ColIndex
            Case IINVNO
                If To_Value(tblInvoice.Columns(ILINE).Text) <> 0 Then
                If Chk_InvNo(tblInvoice.Columns(IINVNO).Text, To_Value(tblInvoice.Columns(ILINE).Text), True) = False Then
                   tblInvoice.Col = IINVNO
                   tblInvoice.SetFocus
                   GoTo Tbl_BeforeColUpdate_Err
                   Exit Sub
                End If
                End If

          
                'unlock the old value if no longer exist in lock table
                If Not IsEmpty(OldValue) Or OldValue <> "" Then
                    If NoMoreInvNo(waInvoice, OldValue, .Bookmark) Then
                        Call RowUnLock(wsConnTime, wsInvKeyType, OldValue, wsFormID)
                    End If
                End If
                
                'Lock the new value
                'if same inv no has been locked, then no need to lock again
                If NoMoreInvNo(waInvoice, .Columns(IINVNO).Text) Then
                    If RowLock(wsConnTime, wsInvKeyType, .Columns(IINVNO).Text, wsFormID, wsUsrId) = False Then
                       gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                       MsgBox gsMsg, vbOKOnly, gsTitle
                    End If
                End If
            
            Case ILINE
                If Not Chk_NoDup(.Row + .FirstRow) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_InvNo(.Columns(IINVNO).Text, .Columns(ILINE).Text, True) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case ICURR
            
                If Chk_Curr(.Columns(ICURR).Text, medDocDate) = False Then
                    gsMsg = "No Such Currency Code!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If getExcRate(.Columns(ICURR).Text, medDocDate.Text, wsExcr, "") = False Then
                    gsMsg = "No Exchange Rate at this period!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    wsExcr = Format(0, gsExrFmt)
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                End If
                
                tblInvoice.Columns(IEXCR) = wsExcr
                Call Calc_InvTotal
                Call Calc_ChargeTotal
                    
                    
            Case ISETAMTORG
            
                If To_Value(.Columns(ISETAMTORG)) < 0 Then
                    .Columns(ISETAMTORG).Text = NBRnd(Abs(.Columns(ISETAMTORG).Text) * -1, giAmtDp)
                End If
                
                If Chk_Amount(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Abs(To_Value(.Columns(ISETAMTORG).Text)) > Abs(To_Value(.Columns(IOSAMT).Text)) Then
                    gsMsg = "Settlement Amount cannot greater than Outstanding Amt!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(ISETAMTORG).Text = NBRnd(.Columns(ISETAMTORG).Text, giAmtDp)
                .Columns(ISETAMTLOC).Text = NBRnd(NBRnd(To_Value(.Columns(ISETAMTORG).Text) * To_Value(.Columns(IEXCR).Text), giAmtDp), giAmtDp)
                Call Calc_InvTotal
                Call Calc_ChargeTotal
                
        End Select
        
    End With
    
    Exit Sub
    
Tbl_BeforeColUpdate_Err:

    tblInvoice.Columns(ColIndex).Text = OldValue
    Cancel = True
    
End Sub
Private Sub tblInvoice_BeforeRowColChange(Cancel As Integer)
    On Error GoTo tblInvoice_BeforeRowColChange_Err
    With tblInvoice
        If .Bookmark <> .DestinationRow Then
            If Chk_InvGrdRow(To_Value(.Bookmark)) = False Then
                Cancel = True
                Exit Sub
            End If
        End If
    End With
    Exit Sub
    
tblInvoice_BeforeRowColChange_Err:
    MsgBox "tblInvoice_BeforeRowColChange_Err!"
    Cancel = True
End Sub

Private Sub tblInvoice_ButtonClick(ByVal ColIndex As Integer)
    
    Dim wsSQL As String
    Dim wiCtr As Integer
    
    On Error GoTo tblInvoice_ButtonClick_Err
    
    With tblInvoice
        Set wcCombo = tblInvoice
        
        Select Case ColIndex
            Case IINVNO
            
                    wsSQL = "SELECT IPHDDOCNO ,"
                    wsSQL = wsSQL & " CONVERT(NVARCHAR(5), REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),IPDTDOCLINE))))  + CONVERT(NVARCHAR(3),IPDTDOCLINE)), "
                    wsSQL = wsSQL & " IPHDJOBNO, "
                    wsSQL = wsSQL & " IPDTINVAMT, "
                    wsSQL = wsSQL & " IPDTBALAMT FROM APIPHD, APIPDT "
                    wsSQL = wsSQL & " WHERE IPHDDOCNO LIKE '%" & IIf(Trim(.SelText) <> "", "", Set_Quote(.Columns(IINVNO).Text)) & "%'"
                    wsSQL = wsSQL & " AND IPHDDOCDATE <= '" & Set_Quote(medDocDate.Text) & "'"
                    wsSQL = wsSQL & " AND IPHDVDRID = " & wlVdrID
                    wsSQL = wsSQL & " AND IPDTBALAMT <> 0  "
                    wsSQL = wsSQL & " AND IPHDSTATUS <> '2' "
                    wsSQL = wsSQL & " AND IPHDDOCID = IPDTDOCID "
                    
                    If waInvoice.UpperBound(1) > -1 Then
                        wsSQL = wsSQL & " AND IPHDDOCNO + CONVERT(NVARCHAR(5), REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),IPDTDOCLINE))))  + CONVERT(NVARCHAR(3),IPDTDOCLINE)) NOT IN ( "
                        For wiCtr = 0 To waInvoice.UpperBound(1)
                            wsSQL = wsSQL & " '" & waInvoice(wiCtr, IINVNO) & waInvoice(wiCtr, ILINE) & IIf(wiCtr = waInvoice.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    wsSQL = wsSQL & " ORDER BY IPHDDOCNO, CONVERT(NVARCHAR(5), REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),IPDTDOCLINE))))  + CONVERT(NVARCHAR(3),IPDTDOCLINE))"

                    Call Ini_Combo(5, wsSQL, tabDetailInfo.Left + .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, tabDetailInfo.Top + .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLINVNO", Me.Width, Me.Height)
               
                
                tblCommon.Visible = True
                tblCommon.SetFocus
        
        Case ICURR
               
                  wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(Len(.Columns(ICURR).Text) > 0, Set_Quote(.Columns(ICURR).Text), "") & "%'"
                  wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Right(wsCtlDte, 2)) & "' "
                  wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Left(wsCtlDte, 4)) & "' "
                  wsSQL = wsSQL & " AND EXCSTATUS = '1' "
                  wsSQL = wsSQL & "ORDER BY EXCCURR "
    
                 Call Ini_Combo(2, wsSQL, tabDetailInfo.Left + .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, tabDetailInfo.Top + .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
     End Select
    End With
    
    Exit Sub
    
tblInvoice_ButtonClick_Err:
    MsgBox "tblInvoice_ButtonClick_Err!"
End Sub

Private Sub tblInvoice_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblInvoice_KeyDown_Err
    
    With tblInvoice
        Select Case KeyCode
        Case vbKeyF4        ' CALL COMBO BOX
            KeyCode = vbDefault
            Call tblInvoice_ButtonClick(.Col)
        
        Case vbKeyF5        ' INSERT LINE
            KeyCode = vbDefault
            If .Bookmark = waInvoice.UpperBound(1) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waInvoice.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
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
                Case IINVNO, ILINE, ICURR, IOSAMT
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case ISETAMTORG, ISETAMTLOC
                    KeyCode = vbKeyDown
                    .Col = IINVNO
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> IINVNO Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> ISETAMTLOC Then
                  .Col = .Col + 1
            End If
            
        End Select
    End With

    Exit Sub
    
tblInvoice_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblInvoice_KeyPress(KeyAscii As Integer)
    Select Case tblInvoice.Col
        
        Case ISETAMTORG
           
            Call Chk_InpNum(KeyAscii, tblInvoice.Text, True, True)
            
        Case ILINE
            Call Chk_InpNum(KeyAscii, tblInvoice.Text, False, False)
        
      
       
    End Select
End Sub



Private Sub tblInvoice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        PopupMenu mnuIPopUp
    End If
    
End Sub

Private Sub tblInvoice_OnAddNew()
    On Error GoTo tblInvoice_OnAddNew_Err
    
    Call Calc_InvTotal
    Call Calc_ChargeTotal
    
    Exit Sub
    
tblInvoice_OnAddNew_Err:

    MsgBox "tblInvoice_OnAddNew_Err!"
    
End Sub

Private Sub tblInvoice_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
    Dim wsCurrDes As String
    Dim chkRow As Variant
    
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblInvoice.Name Then Exit Sub

    With tblInvoice

       If IsEmptyInvRow Then
           .Col = IINVNO
       'Else
       '    cmdInv.Enabled = False
       '    optStlMtd(0).Enabled = False
       '    optStlMtd(1).Enabled = False
       End If


        'if the row is empty, set to the first column
        If Trim(.Columns(IINVNO).Text) <> "" Then
            .Columns(ISETAMTLOC).Text = NBRnd(To_Value(.Columns(ISETAMTORG).Text) * To_Value(.Columns(IEXCR).Text), giAmtDp)
            'Modified May 09 2000 added .update to fix problem
            .Update
        End If
        
        Call Chk_Curr(.Columns(ICURR).Text, medDocDate)
        If .Col = ICURR Then
        End If

        If LastCol = ICURR Or LastCol = ISETAMTORG _
          Or waInvoice.UpperBound(1) = -1 Or waInvoice.UpperBound(1) >= 0 Then
            Call Calc_InvTotal
            Call Calc_ChargeTotal
        End If


    End With
    
    lblDspUnInvAmt = NBRnd(To_Value(LblDspChqAmtLoc) - (To_Value(lblDspInvAmt) + To_Value(lblDspOthAmt) + To_Value(lblDspRemAmtl)), giAmtDp)
    lblDspUnOthAmt = lblDspUnInvAmt
  
    Exit Sub
    
RowColChange_Err:

    MsgBox "RowColChange_Err!"
    
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
        Case tcFind
            Call cmdFind
        Case tcExit
            Unload Me
    End Select
    
End Sub





Private Sub txtChqAmtOrg_GotFocus()
    Call FocusMe(txtChqAmtOrg)
End Sub
Private Sub txtChqAmtOrg_LostFocus()
    If Trim(txtChqAmtOrg.Text) <> "" Then
        txtChqAmtOrg.Text = NBRnd(txtChqAmtOrg, giAmtDp)
    End If
End Sub
Private Sub txtChqAmtOrg_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtChqAmtOrg.Text, False, True)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = vbDefault
       
       If Trim(txtChqAmtOrg.Text) = "" Then
          txtChqAmtOrg = NBRnd("0", giAmtDp)
       End If
       
       If Chk_txtChqAmtOrg(txtChqAmtOrg.Text) = False Then
          Exit Sub
       End If

       
       txtChqAmtOrg.Text = NBRnd(txtChqAmtOrg.Text, giAmtDp)
       
       LblDspRemDte.Caption = medDocDate.Text
       LblDspChqAmtLoc.Caption = NBRnd(To_Value(txtChqAmtOrg.Text) * To_Value(txtExcr.Text), giAmtDp)
       Call Calc_InvTotal
       Call Calc_ChargeTotal
       
       If To_Value(txtChqAmtOrg.Text) = 0 Then
        Call SetFieldStatus("RemFalse")
        txtRemAmt = NBRnd("0", giAmtDp)
        lblDspRemAmtl.Caption = NBRnd("0", giAmtDp)
        LblDspRemDte.Caption = ""
        cboTMPML.Text = ""
        tabDetailInfo.Tab = 0
        If Chk_KeyFld Then
        tblInvoice.SetFocus
        End If
       Else
        Call SetFieldStatus("RemTrue")
        txtRemAmt.SetFocus
       End If
       
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
                cboMLCode.SetFocus
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



Private Sub cboVdrCode_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrCode
    
    If gsLangID = "1" Then
        wsSQL = "SELECT VDRCODE, VDRNAME FROM mstVENDOR "
        wsSQL = wsSQL & "WHERE VDRCODE LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSQL = wsSQL & "AND VDRSTATUS = '1' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY VDRCODE "
    Else
        wsSQL = "SELECT VDRCODE, VDRNAME FROM mstVENDOR "
        wsSQL = wsSQL & "WHERE VDRCODE LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSQL = wsSQL & "AND VDRSTATUS = '1' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY VDRCODE "
    End If
    Call Ini_Combo(2, wsSQL, cboVdrCode.Left, cboVdrCode.Top + cboVdrCode.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    
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
        gsMsg = "必需輸入客戶編碼!"
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
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
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
    wsSQL = wsSQL & "FROM  mstVENDOR "
    wsSQL = wsSQL & "WHERE VDRID = " & wlVdrID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "VDRCURR")
          Else
        cboCurr.Text = ""
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
    
    
   
    
    'get Due Date Payment Term

End Sub



Private Sub Ini_Grid()
    
    Dim wiCtr As Integer
    
        With tblInvoice
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = IINVNO To IDummy
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case IINVNO
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 15
                Case ILINE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 3
                Case ICURR
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 3
                Case IEXCR
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case IOSAMT
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case ISETAMTORG
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case ISETAMTLOC
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case IIPDTID
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case IDummy
                    .Columns(wiCtr).DataWidth = 0
                    .Columns(wiCtr).Locked = True
                    
            End Select
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With

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
        
        For wiCtr = OMLCODE To ODummy
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case OMLCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                Case OCURR
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 3
                Case OEXCR
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsExrFmt
                Case OOTHAMTORG
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case OOTHAMTLOC
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case ODummy
                    .Columns(wiCtr).DataWidth = 0
                    .Columns(wiCtr).Locked = True
                    
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
Dim wsDes As String
Dim wsExcr As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case OMLCODE
                
                If Chk_grdMLClass(.Columns(ColIndex).Text, "G", wsDes) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Trim(.Columns(OCURR)) = "" Then
                   .Columns(OCURR) = wsBaseCurCd
                   .Columns(OEXCR) = wsBaseExcr
                End If
                Call chk_BaseCurr
                
            Case OCURR
                
                If Chk_Curr(.Columns(OCURR).Text, medDocDate) = False Then
                    gsMsg = "No Such Currency Code!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If getExcRate(.Columns(OCURR).Text, medDocDate.Text, wsExcr, "") = False Then
                    gsMsg = "No Exchange Rate at this period!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    wsExcr = Format(0, gsExrFmt)
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                End If
                
                
                .Columns(OEXCR).Text = wsExcr
                Call chk_BaseCurr
                
           Case OEXCR
                If To_Value(.Columns(ColIndex).Text) > "99999.9999" Or _
                    To_Value(.Columns(ColIndex).Text) <= 0 Then
                    
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(OOTHAMTLOC).Text = NBRnd(To_Value(.Columns(OOTHAMTORG)) * _
                    To_Value(.Columns(OEXCR)), giAmtDp)
                Call Calc_ChargeTotal
                Call Calc_InvTotal
                
            Case OOTHAMTORG
                                
                If Chk_Amount(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(OOTHAMTORG).Text = NBRnd(.Columns(OOTHAMTORG).Text, giAmtDp)
                .Columns(OOTHAMTLOC).Text = NBRnd(To_Value(.Columns(OOTHAMTORG)) * _
                    To_Value(.Columns(OEXCR)), giAmtDp)
                Call Calc_ChargeTotal
                Call Calc_InvTotal
                
                
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
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case OMLCODE
                
                wsSQL = "SELECT MLCODE, MLDESC FROM mstMerchClass "
                wsSQL = wsSQL & " WHERE MLSTATUS <> '2' "
                wsSQL = wsSQL & " AND MLCODE LIKE '%" & Set_Quote(.Columns(OMLCODE).Text) & "%' "
                wsSQL = wsSQL & " AND MLTYPE = 'G' "
                wsSQL = wsSQL & " ORDER BY MLCODE "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case OCURR
                
            
               wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(Len(.Columns(OCURR).Text) > 0, Set_Quote(.Columns(OCURR).Text), "") & "%'"
               wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Right(wsCtlDte, 2)) & "' "
               wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Left(wsCtlDte, 4)) & "' "
               wsSQL = wsSQL & " AND EXCSTATUS = '1' "
               wsSQL = wsSQL & "ORDER BY EXCCURR "
                  
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
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
                Case OMLCODE, OEXCR
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case OOTHAMTORG, OOTHAMTLOC
                    KeyCode = vbKeyDown
                    .Col = OMLCODE
                Case OCURR
                     KeyCode = vbDefault
                    If UCase(Trim(.Columns(OCURR).Text)) = UCase(wsBaseCurCd) Then
                        .Col = OOTHAMTORG
                    Else
                        .Col = OEXCR
                    End If
            End Select
        Case vbKeyLeft
        
            Select Case .Col
                Case OCURR, OEXCR, OOTHAMTLOC
                    KeyCode = vbDefault
                    .Col = .Col - 1
                Case OOTHAMTORG
                     KeyCode = vbDefault
                    If UCase(Trim(.Columns(OCURR).Text)) = UCase(wsBaseCurCd) Then
                        .Col = OCURR
                    Else
                        .Col = OEXCR
                    End If
            End Select
        
        Case vbKeyRight
            
            Select Case .Col
                Case OMLCODE, OEXCR, OOTHAMTORG
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case OCURR
                     KeyCode = vbDefault
                    If UCase(Trim(.Columns(OCURR).Text)) = UCase(wsBaseCurCd) Then
                        .Col = OOTHAMTORG
                    Else
                        .Col = OEXCR
                    End If
            End Select
        
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
    
        Case OOTHAMTORG
            
            Call Chk_InpNum(KeyAscii, tblDetail.Text, True, True)
        
        Case OEXCR
            
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
        
      
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = OMLCODE
        End If
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case OMLCODE
                    If Trim(.Columns(.Col).Text) <> "" Then
                    Call Chk_grdMLClass(.Columns(.Col).Text, "G", "")
                    End If
      
                Case OCURR
                
                If Chk_Curr(.Columns(.Col).Text, medDocDate) = False Then
                    gsMsg = "No Such Currency Code!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
                
                
                
            End Select
        End If
        Call Calc_ChargeTotal
        Call Calc_InvTotal
        If Trim(.Columns(OMLCODE).Text) <> "" Then .Columns(OOTHAMTLOC).Text = NBRnd(To_Value(.Columns(OOTHAMTORG)) * To_Value(.Columns(OEXCR)), giAmtDp)

    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub



Private Function Chk_grdMLCode(inNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdMLCode = False
    
    If Trim(inNo) = "" Then
        Chk_grdMLCode = True
        Exit Function
    End If
    
    wsSQL = "SELECT *  FROM mstMerchClass"
    wsSQL = wsSQL & " WHERE MLCode = '" & Set_Quote(inNo) & "' "
    wsSQL = wsSQL & " AND MLTYPE = 'G' "
    wsSQL = wsSQL & " AND MLSTATUS = '1' "
       
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_grdMLCode = True
        

End Function








Private Function Chk_Amount(inAmt As String) As Integer
    
    Chk_Amount = False
    
    If Trim(inAmt) = "" Then
        gsMsg = "必需輸入金額!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
    If To_Value(inAmt) = 0 Then
        gsMsg = "數量必需大於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If To_Value(inAmt) > gsMaxVal Then
        gsMsg = "數量太大!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_Amount = True

End Function

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(OMLCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, OMLCODE)) = "" And _
                   Trim(waResult(inRow, OCURR)) = "" And _
                   Trim(waResult(inRow, OEXCR)) = "" And _
                   Trim(waResult(inRow, OOTHAMTORG)) = "" And _
                   Trim(waResult(inRow, OOTHAMTLOC)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
End Function


Private Function IsEmptyInvRow(Optional inRow) As Boolean

    IsEmptyInvRow = True
    
        If IsMissing(inRow) Then
            With tblInvoice
                If Trim(.Columns(IINVNO)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waInvoice.UpperBound(1) >= 0 Then
                If Trim(waInvoice(inRow, IINVNO)) = "" And _
                   Trim(waInvoice(inRow, ILINE)) = "" And _
                   Trim(waInvoice(inRow, ICURR)) = "" And _
                   Trim(waInvoice(inRow, IOSAMT)) = "" And _
                   Trim(waInvoice(inRow, ISETAMTORG)) = "" And _
                   Trim(waInvoice(inRow, ISETAMTLOC)) = "" Then
                   Exit Function
                End If
            End If
        End If
        
    IsEmptyInvRow = False
    
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
        
        If Chk_MLClass(waResult(LastRow, OMLCODE), "G", wsDes) = False Then
            .Col = OMLCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_Curr(waResult(LastRow, OCURR), medDocDate) = False Then
                .Col = OCURR
                .Row = LastRow
                Exit Function
        End If
        
        If To_Value(waResult(LastRow, OEXCR)) > "99999.9999" Or _
                    To_Value(waResult(LastRow, OEXCR)) <= 0 Then
                .Col = OEXCR
                .Row = LastRow
                Exit Function
        End If
                
        
        If Chk_Amount(waResult(LastRow, OOTHAMTORG)) = False Then
            .Col = OOTHAMTORG
            .Row = LastRow
            Exit Function
        End If
        
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function


Private Sub Calc_InvTotal()

    Dim wdStlTot As Double
    Dim wdStlTotL As Double
    Dim wlCtr As Long
    
    On Error GoTo Calc_InvTotal_Err
    
    wdStlTot = 0
    wdStlTotL = 0
    
    wlCtr = To_Value(tblInvoice.FirstRow) + tblInvoice.Row
    
    If waInvoice.UpperBound(1) = -1 Then
        lblDspInvAmt.Caption = NBRnd("0", giAmtDp)
        lblDspUnInvAmt.Caption = NBRnd("0", giAmtDp)
        Exit Sub
    End If

    If wlCtr > waInvoice.UpperBound(1) Then Exit Sub
    
    If Not IsEmptyInvRow(wlCtr) Then
       waInvoice(wlCtr, ISETAMTORG) = tblInvoice.Columns(ISETAMTORG).Text
       waInvoice(wlCtr, ISETAMTLOC) = tblInvoice.Columns(ISETAMTLOC).Text
       waInvoice(wlCtr, IEXCR) = tblInvoice.Columns(IEXCR).Text
    End If
    
    For wlCtr = 0 To waInvoice.UpperBound(1)
        wdStlTot = NBRnd(wdStlTot + (To_Value(waInvoice(wlCtr, ISETAMTLOC)) / IIf(To_Value(waInvoice(wlCtr, IEXCR)) = 0, 1, To_Value(waInvoice(wlCtr, IEXCR)))), giAmtDp)
        wdStlTotL = NBRnd(wdStlTotL + (To_Value(waInvoice(wlCtr, ISETAMTLOC))), giAmtDp)
    Next
    
    lblDspInvAmt.Caption = NBRnd(wdStlTotL, giAmtDp)
    lblDspUnInvAmt = NBRnd(To_Value(LblDspChqAmtLoc) - (To_Value(lblDspInvAmt) + To_Value(lblDspOthAmt) + To_Value(lblDspRemAmtl)), giAmtDp)
    lblDspUnOthAmt = lblDspUnInvAmt

    'Modified 09 May to remove problem of not appearing in the grid
    tblInvoice.Update
    
    Exit Sub
    
Calc_InvTotal_Err:
      MsgBox "Calc_InvTotal"
    
End Sub




Private Function cmdDel() As Boolean

    Dim wsGenDte As String
    Dim wsDteTim As String
    
    Dim adcmdDelete As New ADODB.Command
    Dim wsDocNo As String
    Dim i As Integer
    
    cmdDel = False
    
    MousePointer = vbHourglass
    
    On Error GoTo cmdDelete_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    wsDteTim = Change_SQLDate(Now)
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbReadOnly Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
        'check each detail line
    If Chk_LockedRecords Then
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
        
    adcmdDelete.CommandText = "USP_AP003A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wlKey)
    Call SetSPPara(adcmdDelete, 3, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 4, wlVdrID)
    Call SetSPPara(adcmdDelete, 5, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 6, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 7, txtExcr.Text)
    Call SetSPPara(adcmdDelete, 8, To_Value(txtChqAmtOrg.Text))
    Call SetSPPara(adcmdDelete, 9, NBRnd(To_Value(txtExcr.Text) * To_Value(txtChqAmtOrg.Text), giAmtDp))
    Call SetSPPara(adcmdDelete, 10, cboMLCode.Text)
    Call SetSPPara(adcmdDelete, 11, "")
    Call SetSPPara(adcmdDelete, 12, txtRmk.Text)
    Call SetSPPara(adcmdDelete, 13, wsFormID)
    Call SetSPPara(adcmdDelete, 14, gsUserID)
    Call SetSPPara(adcmdDelete, 15, wsGenDte)
    Call SetSPPara(adcmdDelete, 16, wsDteTim)
    
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 17)
    wsDocNo = GetSPPara(adcmdDelete, 18)
    
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
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.cboVdrCode.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.txtRmk.Enabled = False
            Me.txtChqAmtOrg.Enabled = False
            
            
            Me.tblInvoice.Enabled = False
            Me.tblDetail.Enabled = False
            
        Case "AfrActEdit"
        
            Me.cboDocNo.Enabled = True
            
            
        Case "AfrKeyAdd"
        
            Me.cboDocNo.Enabled = False
            Me.cboVdrCode.Enabled = True
       
       Case "AfrKeyEdit"
       
            Me.cboDocNo.Enabled = False
            Me.cboVdrCode.Enabled = False
        
        Case "AfrKey"
        '    Me.cboDocNo.Enabled = False
        '    Me.cboVdrCode.Enabled = True
            
            
            Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.cboMLCode.Enabled = True
            Me.txtRmk.Enabled = True
            Me.txtChqAmtOrg.Enabled = True
            
            
            If wiAction <> AddRec Then
                Me.tblDetail.Enabled = True
                Me.tblInvoice.Enabled = True
            End If
            
         Case "RemFalse"
            Me.cboTMPML.Enabled = False
            Me.txtRemAmt.Enabled = False
            
         Case "RemTrue"
            Me.cboTMPML.Enabled = True
            Me.txtRemAmt.Enabled = True
                        
            
    End Select
End Sub

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
    
        .TableID = wsKeyType
        .TableType = wsSrcCd
        .TableKey = "APSHDocNo"
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
    vFilterAry(1, 2) = "APSHDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "APSHDocDate"
    
    vFilterAry(3, 1) = "Vendor #"
    vFilterAry(3, 2) = "VdrCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "APSHDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "APSHDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Vendor#"
    vAry(3, 2) = "VdrCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Vendor Name"
    vAry(4, 2) = "VdrName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT APSHDocNo, APSHDocDate, mstVendor.VdrCode,  mstVendor.VdrName "
        wsSQL = wsSQL + "FROM MstVendor, ArStHd "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE APSHStatus = '1' And APSHVdrID = VdrID "
        .sBindOrderSQL = "ORDER BY APSHDocNo"
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
        
        
        txtRmk.SetFocus
       
    End If
    
End Sub

Private Sub cboMLCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass "
    wsSQL = wsSQL & " WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSQL = wsSQL & " AND MLSTATUS = '1' "
    wsSQL = wsSQL & " AND MLTYPE = 'B' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboMLCode.Left, cboMLCode.Top + cboMLCode.Height, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
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
        cboMLCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_MLClass(cboMLCode, "B", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboMLCode.SetFocus
        lblDspMLDesc = ""
       Exit Function
    End If
    
    lblDspMLDesc = wsDesc
    
    Chk_cboMLCode = True
    
End Function





Private Sub cboTmpML_GotFocus()
    FocusMe cboTMPML
End Sub

Private Sub cboTmpML_LostFocus()
    FocusMe cboTMPML, True
End Sub


Private Sub cboTmpML_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboTMPML, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboTmpML = False Then
                Exit Sub
        End If
        
        
        tabDetailInfo.Tab = 0
        If Chk_KeyFld Then
        tblInvoice.SetFocus
        End If
       
    End If
    
End Sub

Private Sub cboTmpML_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboTMPML
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass "
    wsSQL = wsSQL & " WHERE MLCode LIKE '%" & IIf(cboTMPML.SelLength > 0, "", Set_Quote(cboTMPML.Text)) & "%' "
    wsSQL = wsSQL & " AND MLSTATUS = '1' "
    wsSQL = wsSQL & " AND MLTYPE = 'R' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboTMPML.Left, cboTMPML.Top + cboTMPML.Height, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboTmpML() As Boolean
Dim wsDesc As String

    Chk_cboTmpML = False
     
    If To_Value(txtRemAmt) = 0 Then
        cboTMPML.Text = ""
        Chk_cboTmpML = True
        Exit Function
    End If
     
    If Trim(cboTMPML.Text) = "" Then
        gsMsg = "必需輸入會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboTMPML.SetFocus
        Exit Function
    End If
    
    
    If Chk_MLClass(cboTMPML, "R", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboTMPML.SetFocus
        lblDspTmpMlDesc = ""
       Exit Function
    End If
    
    lblDspTmpMlDesc = wsDesc
    
    Chk_cboTmpML = True
    
End Function






Private Sub txtRemAmt_GotFocus()
    FocusMe txtRemAmt
End Sub

Private Sub txtRemAmt_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtRemAmt.Text, True, True)
    
    If KeyAscii = vbKeyReturn Then
       KeyAscii = vbDefault
       
       If Trim(txtRemAmt.Text) = "" Then
          txtRemAmt.Text = NBRnd("0", giAmtDp)
          cboTMPML.SetFocus
          Exit Sub
       End If
       
        If Chk_txtRemAmt(txtRemAmt.Text) = False Then
          Exit Sub
       End If
       
        lblDspRemAmtl.Caption = Format(NBRnd(To_Value(txtRemAmt.Text) * To_Value(txtExcr.Text), giAmtDp), gsAmtFmt)
        txtRemAmt.Text = NBRnd(To_Value(txtRemAmt.Text), giAmtDp)
        lblDspUnInvAmt = NBRnd(To_Value(LblDspChqAmtLoc) - (To_Value(lblDspInvAmt) + To_Value(lblDspOthAmt) + To_Value(lblDspRemAmtl)), giAmtDp)
        lblDspUnOthAmt = lblDspUnInvAmt
  
        cboTMPML.SetFocus
    End If
    
End Sub


Private Sub txtRemAmt_LostFocus()
 FocusMe txtRemAmt, True
 If Trim(txtRemAmt.Text) <> "" Then
        txtRemAmt.Text = NBRnd(txtRemAmt, giAmtDp)
 End If
 
End Sub

Private Sub txtRmk_GotFocus()

        
        FocusMe txtRmk

End Sub

Private Sub txtRmk_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtRmk, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        txtChqAmtOrg.SetFocus
        
        
    End If
End Sub

Private Sub txtRmk_LostFocus()
        
        FocusMe txtRmk, True

End Sub









Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuOPopUp
    End If
    

End Sub

Private Sub mnuOPopUpSub_Click(Index As Integer)
    Call Call_OPopUpMenu(waPopUpSub, Index)
End Sub

Private Sub mnuIPopUpSub_Click(Index As Integer)
    Call Call_IPopUpMenu(waPopUpSub, Index)
End Sub

Private Sub Call_IPopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

    Dim wsAct As String
    
    wsAct = inArray(inMnuIdx, 0)
    
    With tblInvoice
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

Private Sub Call_OPopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

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
            
            If .Bookmark = waInvoice.UpperBound(1) Then Exit Sub
            If IsEmptyInvRow Then Exit Sub
            waInvoice.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
             
    
End Sub
Private Function Chk_DocNo(ByVal InDocNo As String, ByRef OutStatus As String, ByRef OutUpdFlg As String, ByRef OutPgmNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    OutStatus = ""
    OutUpdFlg = ""
    Chk_DocNo = False
    
    wsSQL = "SELECT APSHSTATUS, APSHUPDFLG, APSHPGMNO FROM APSTHD "
    wsSQL = wsSQL & " WHERE APSHDOCNO = '" & Set_Quote(InDocNo) & "' "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount <= 0 Then
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    OutStatus = ReadRs(rsRcd, "APSHSTATUS")
    OutUpdFlg = ReadRs(rsRcd, "APSHUPDFLG")
    OutPgmNo = ReadRs(rsRcd, "APSHPGMNO")
    
    rsRcd.Close
    Set rsRcd = Nothing
    Chk_DocNo = True
    
    
    

End Function



Private Function Chk_VdrCode(ByVal InVdrNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT VdrID, VdrName, VdrTel, VdrFax FROM mstVendor WHERE VdrCode = '" & Set_Quote(InVdrNo) & "' "
    wsSQL = wsSQL & "And VdrStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "VdrID")
        OutName = ReadRs(rsRcd, "VdrName")
        OutTel = ReadRs(rsRcd, "VdrTel")
        OutFax = ReadRs(rsRcd, "VdrFax")
        Chk_VdrCode = True
        
    Else
    
        OutID = 0
        OutName = ""
        OutTel = ""
        OutFax = ""
        Chk_VdrCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Private Function Chk_txtChqAmtOrg(inAmt As String) As Integer
    
    Chk_txtChqAmtOrg = False
    
   
    
    If To_Value(inAmt) > gsMaxVal Then
        gsMsg = "數量太大!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtChqAmtOrg.SetFocus
        Exit Function
    End If
    
    Chk_txtChqAmtOrg = True

End Function

Private Function Chk_txtRemAmt(inAmt As String) As Integer
    
    Chk_txtRemAmt = False
    
   
    
    If To_Value(inAmt) > gsMaxVal Then
        gsMsg = "數量太大!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtRemAmt.SetFocus
        Exit Function
    End If
    
    Chk_txtRemAmt = True

End Function

Private Sub Calc_ChargeTotal()

    Dim wdStlTot As Double
    Dim wdStlTotL As Double
    Dim wlCtr As Long
    
    On Error GoTo Calc_ChargeTotal_Err
    
    wdStlTot = 0
    wdStlTotL = 0
    
   wlCtr = To_Value(tblDetail.FirstRow) + tblDetail.Row
    
    If Abs(wlCtr) > waResult.UpperBound(1) Then
        lblDspUnInvAmt = NBRnd(To_Value(LblDspChqAmtLoc) - (To_Value(lblDspInvAmt) + To_Value(lblDspOthAmt) + To_Value(lblDspRemAmtl)), giAmtDp)
        lblDspUnOthAmt = lblDspUnInvAmt
  
        If waResult.UpperBound(1) = -1 Then
            lblDspOthAmt.Caption = NBRnd("0", giAmtDp)
            lblDspUnOthAmt.Caption = NBRnd("0", giAmtDp)
            Exit Sub
        End If
        Exit Sub
    End If
    
    If waResult.UpperBound(1) = -1 Then
       lblDspOthAmt.Caption = NBRnd("0", giAmtDp)
       lblDspUnOthAmt.Caption = NBRnd("0", giAmtDp)
       Exit Sub
    End If
    
    If Not IsEmptyRow(wlCtr) Then
       waResult(wlCtr, OOTHAMTORG) = tblDetail.Columns(OOTHAMTORG).Text
       waResult(wlCtr, OOTHAMTLOC) = NBRnd(To_Value(tblDetail.Columns(OOTHAMTORG).Text) * To_Value(tblDetail.Columns(OEXCR).Text), giAmtDp)
    End If
    
    For wlCtr = 0 To waResult.UpperBound(1)
        wdStlTot = NBRnd(wdStlTot + To_Value(waResult(wlCtr, OOTHAMTORG)), giAmtDp)
        wdStlTotL = NBRnd(wdStlTotL + To_Value(waResult(wlCtr, OOTHAMTLOC)), giAmtDp)
    Next
    
    lblDspOthAmt.Caption = NBRnd(wdStlTotL, giAmtDp)
    lblDspUnInvAmt = NBRnd(To_Value(LblDspChqAmtLoc) - (To_Value(lblDspInvAmt) + To_Value(lblDspOthAmt) + To_Value(lblDspRemAmtl)), giAmtDp)
    lblDspUnOthAmt = lblDspUnInvAmt
  
    Exit Sub
    
Calc_ChargeTotal_Err:
   MsgBox "Calc_ChargeTotal_Err"

End Sub

Private Function Chk_InvNo(ByVal inInvNo As String, ByVal inInvLn As String, Optional inDtl As Boolean) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsExcRat As String
    Dim wlCtr As Long
    
    On Error GoTo Chk_InvNo_Err
    
    Chk_InvNo = False
    
    
    If IsMissing(inDtl) Then inDtl = False
    
        wsSQL = "SELECT IPDTID, IPHDDOCNO , REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),IPDTDOCLINE))))  + CONVERT(NVARCHAR(3),IPDTDOCLINE) AS LN, "
        wsSQL = wsSQL & " IPHDCURR, IPDTBALAMT, IPDTBALAMTL, IPHDEXCR FROM APIPHD, APIPDT "
        wsSQL = wsSQL & " WHERE IPHDDOCNO = '" & Set_Quote(inInvNo) & "'"
        wsSQL = wsSQL & " AND  REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),IPDTDOCLINE))))  + CONVERT(NVARCHAR(3),IPDTDOCLINE) = '" & Set_Quote(Format(inInvLn, "00#")) & "'"
        wsSQL = wsSQL & " AND IPHDVDRID = " & wlVdrID
        wsSQL = wsSQL & " AND IPHDDOCID = IPDTDOCID AND IPDTBALAMT <> 0"
        wsSQL = wsSQL & " AND IPHDSTATUS <> '2'"
        
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount = 0 Then
          
        gsMsg = "No Such Invoice No!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       
       rsRcd.Close
       Set rsRcd = Nothing
       Exit Function
    End If
    
    Chk_InvNo = True
    
    If Not inDtl Then
       rsRcd.Close
       Set rsRcd = Nothing
       Exit Function
    End If

    
     With tblInvoice
                .Columns(IINVNO).Text = ReadRs(rsRcd, "IPHDDOCNO")
                .Columns(ILINE).Text = ReadRs(rsRcd, "LN")
                .Columns(ICURR).Text = ReadRs(rsRcd, "IPHDCURR")
                .Columns(IOSAMT).Text = NBRnd(ReadRs(rsRcd, "IPDTBALAMT"), giAmtDp)
                .Columns(ISETAMTORG).Text = NBRnd(ReadRs(rsRcd, "IPDTBALAMT"), giAmtDp)
                .Columns(ISETAMTLOC).Text = NBRnd(ReadRs(rsRcd, "IPDTBALAMTL"), giAmtDp)
                .Columns(IEXCR).Text = ReadRs(rsRcd, "IPHDEXCR")
                .Columns(IIPDTID).Text = ReadRs(rsRcd, "IPDTID")
                
    
                
            
            Call Chk_Curr(.Columns(ICURR).Text, medDocDate)
            'tblInvoice.Columns(Tab1StlAmt).Locked = False
            
    End With
    
           
           
    Chk_InvNo = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Exit Function
    
Chk_InvNo_Err:
    MsgBox "Chk_InvNo_Err!"
    
End Function
Private Function NoMoreInvNo(ByVal inArray As XArrayDB, ByVal inInvNo As String, _
    Optional inLnNo As Variant) As Boolean
    'contents of inarray before new value is updated
    'content of inInvNo is the new inv no entered
    
    On Error GoTo NoMoreInvNo_Err
    Dim wiRowCtr As Integer
    NoMoreInvNo = True
    For wiRowCtr = 0 To inArray.UpperBound(1)
        If IsMissing(inLnNo) Then
            If inInvNo = inArray(wiRowCtr, IINVNO) Then
                NoMoreInvNo = False
                Exit Function
            End If
        Else
            If inInvNo = inArray(wiRowCtr, IINVNO) And wiRowCtr <> inLnNo Then
                NoMoreInvNo = False
                Exit Function
            Else
                If inInvNo <> inArray(wiRowCtr, IINVNO) And inArray.UpperBound(1) = 0 Then
                    NoMoreInvNo = False
                    Exit Function
                End If
            End If
            
        End If
    Next wiRowCtr
    Exit Function
    
NoMoreInvNo_Err:
   MsgBox "NoMoreInvNo_Err!"
    
End Function

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoDup = False
    
    wsCurRec = tblInvoice.Columns(IINVNO)
    wsCurRecLn = Format(tblInvoice.Columns(ILINE), "00#")
   
        For wlCtr = 0 To waInvoice.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waInvoice(wlCtr, IINVNO) And _
                  wsCurRecLn = waInvoice(wlCtr, ILINE) Then
                  gsMsg = "Duplicate Invoice Line !"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    
    Chk_NoDup = True

End Function
Private Function Chk_InvGrdRow(ByVal LastRow As Long) As Boolean

    Chk_InvGrdRow = False
    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    
    'added 09 May
    If To_Value(LastRow) > waInvoice.UpperBound(1) Then
       Chk_InvGrdRow = True
       Exit Function
    End If
    
    If IsEmptyInvRow(To_Value(LastRow)) = True Then
       Chk_InvGrdRow = True
       Exit Function
    End If
    
    With tblInvoice
        'added 09 May
        'If Chk_InvNo(waInvoice(wlCtr, Tab1InvNo), waInvoice(wlCtr, Tab1InvLn), False, True) = False Then
        '    .Col = Tab1InvNo
        '    Exit Function
        'End If
        
        'added 09 May
        
         If Chk_Curr(waInvoice(LastRow, ICURR), medDocDate) = False Then
               gsMsg = "No Such Currency Code!"
               MsgBox gsMsg, vbOKOnly, gsTitle
               .Col = ICURR
               Exit Function
        End If
                
        
        If Chk_Amount(waInvoice(LastRow, ISETAMTORG)) = False Then
            .Col = ISETAMTORG
            Exit Function
        End If
         
        If Abs(To_Value(waInvoice(LastRow, ISETAMTORG))) > Abs(To_Value(waInvoice(LastRow, IOSAMT))) Then
            gsMsg = "Settlement Amount cannot greater than Outstanding Amt!"
            MsgBox gsMsg, vbOKOnly, gsTitle
           .Col = ISETAMTORG
           Exit Function
        End If
       
    End With
    Chk_InvGrdRow = True
    
End Function

Private Sub chk_BaseCurr()

     With tblDetail
     
     If UCase(Trim(wsBaseCurCd)) = UCase(tblDetail.Columns(OCURR).Text) Then
        .Columns(OEXCR).Text = NBRnd("1", giExrDp)
        .Columns(OEXCR).Locked = True
     Else
        .Columns(OEXCR).Locked = False
     End If
      
     End With
End Sub


Private Function Chk_LockedRecords() As Boolean
    Dim wiRowCtr As Integer
    Chk_LockedRecords = True
    
    For wiRowCtr = 0 To waInvoice.UpperBound(1)
        If waInvoice(wiRowCtr, IINVNO) <> "" Then
            If ReadOnlyMode(wsConnTime, wsInvKeyType, waInvoice(wiRowCtr, IINVNO), wsFormID) Then
                gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
                MsgBox gsMsg, vbOKOnly, gsTitle
                tabDetailInfo.Tab = 0
                tblInvoice.SetFocus
                tblInvoice.Bookmark = wiRowCtr
                tblInvoice.Col = IINVNO
                MousePointer = vbDefault
                Exit Function
            End If
        End If
    Next wiRowCtr

    Chk_LockedRecords = False
End Function
Public Function Chk_grdMLClass(ByVal inCode As String, ByVal inType As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT MLDesc FROM mstMerchClass WHERE MLCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And MLStatus = '1' "
    wsSQL = wsSQL & "And MLType = '" & inType & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "MLDesc")
        Chk_grdMLClass = True
        
    Else
        gsMsg = "No Such A/C Class!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        OutDesc = ""
        Chk_grdMLClass = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function
