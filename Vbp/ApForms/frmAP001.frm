VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAP001 
   Caption         =   "訂貨單"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "frmAP001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "frmAP001.frx":030A
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboJobNo 
      Height          =   300
      Left            =   4800
      TabIndex        =   43
      Top             =   840
      Width           =   2055
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
      Left            =   7560
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
            Picture         =   "frmAP001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAP001.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   22
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
      Height          =   4455
      Left            =   0
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Header Information"
      TabPicture(0)   =   "frmAP001.frx":66AD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picRmk"
      Tab(0).Control(1)=   "cboRmkCode"
      Tab(0).Control(2)=   "cboPayCode"
      Tab(0).Control(3)=   "cboMLCode"
      Tab(0).Control(4)=   "medDueDate"
      Tab(0).Control(5)=   "lblRmk"
      Tab(0).Control(6)=   "lblRmkCode"
      Tab(0).Control(7)=   "lblMlCode"
      Tab(0).Control(8)=   "lblDspMLDesc"
      Tab(0).Control(9)=   "lblPayCode"
      Tab(0).Control(10)=   "lblDspPayDesc"
      Tab(0).Control(11)=   "lblDueDate"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmAP001.frx":66C9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblInvAmtLoc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDspInvAmtLoc"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDspInvAmtOrg"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblInvAmtOrg"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "tblDetail"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.PictureBox picRmk 
         BackColor       =   &H80000009&
         Height          =   1815
         Left            =   -73200
         ScaleHeight     =   1755
         ScaleWidth      =   7635
         TabIndex        =   41
         Top             =   1680
         Width           =   7695
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   14
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1395
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   13
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1035
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   690
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   0
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   345
            Width           =   7545
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   9
         Top             =   1260
         Width           =   1890
      End
      Begin VB.ComboBox cboPayCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   6
         Top             =   180
         Width           =   2370
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   7
         Top             =   540
         Width           =   2370
      End
      Begin MSMask.MaskEdBox medDueDate 
         Height          =   285
         Left            =   -73200
         TabIndex        =   8
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "frmAP001.frx":66E5
         TabIndex        =   15
         Top             =   120
         Width           =   9495
      End
      Begin VB.Label lblRmk 
         Caption         =   "RMK"
         Height          =   240
         Left            =   -74760
         TabIndex        =   42
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label lblRmkCode 
         Caption         =   "RMKCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   40
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label lblInvAmtOrg 
         Caption         =   "NETAMTORG"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   3660
         Width           =   1875
      End
      Begin VB.Label lblDspInvAmtOrg 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5040
         TabIndex        =   38
         Top             =   3660
         Width           =   1290
      End
      Begin VB.Label lblDspInvAmtLoc 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   8280
         TabIndex        =   37
         Top             =   3660
         Width           =   1290
      End
      Begin VB.Label lblInvAmtLoc 
         Caption         =   "NETAMTLOC"
         Height          =   255
         Left            =   6480
         TabIndex        =   36
         Top             =   3660
         Width           =   1755
      End
      Begin VB.Label lblMlCode 
         Caption         =   "MLCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   35
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblDspMLDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   34
         Top             =   540
         Width           =   5175
      End
      Begin VB.Label lblPayCode 
         Caption         =   "PAYCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   33
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblDspPayDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   32
         Top             =   180
         Width           =   5175
      End
      Begin VB.Label lblDueDate 
         Caption         =   "DUEDATE"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   960
         Width           =   1545
      End
   End
   Begin VB.Label lblJobNo 
      Caption         =   "CUSCODE"
      Height          =   255
      Left            =   3360
      TabIndex        =   44
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblVdrTel 
      Caption         =   "CUSTEL"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblVdrFax 
      Caption         =   "CUSFAX"
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblDspVdrFax 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4800
      TabIndex        =   27
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblVdrName 
      Caption         =   "CUSNAME"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lblDspVdrTel 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   25
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblExcr 
      Caption         =   "EXCR"
      Height          =   255
      Left            =   7000
      TabIndex        =   24
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label LblCurr 
      Caption         =   "CURR"
      Height          =   255
      Left            =   7000
      TabIndex        =   23
      Top             =   1260
      Width           =   1200
   End
   Begin VB.Label lblDspVdrName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   20
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lblVdrCode 
      Caption         =   "CUSCODE"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lblDocDate 
      Caption         =   "DOCDATE"
      Height          =   255
      Left            =   7000
      TabIndex        =   18
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label lblRevNo 
      Caption         =   "REVNO"
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   540
      Width           =   1335
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
      TabIndex        =   16
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
Attribute VB_Name = "frmAP001"
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
Private wbReadOnly As Boolean



Private Const GMLCODE = 0
Private Const GDESC = 1
Private Const GJOBNO = 2
Private Const GCUSPO = 3
Private Const GAMT = 4
Private Const GDTLID = 5
Private Const GDummy = 6


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


Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "APIPHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsSrcCd As String

Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String
Private wsSOPFlg As String



Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, GMLCODE, GDTLID
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

      
    
    wsOldVdrNo = ""
    wsOldCurCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""

    
    wlKey = 0
    wlVdrID = 0
    wbReadOnly = False
    
    
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
            cboPayCode.SetFocus
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
  
    
    wsSQL = "SELECT IPHDDOCNO, IPHDJOBNO, IPHDDOCDATE, IPHDCURR, SUM(IPDTINVAMT), SUM(IPDTBALAMT) "
    wsSQL = wsSQL & " FROM APIPHD, APIPDT "
    wsSQL = wsSQL & " WHERE IPHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND IPHDDOCID  = IPDTDOCID "
    wsSQL = wsSQL & " AND IPHDTRNCODE  = '" & Set_Quote(wsTrnCd) & "' "
 '   wsSql = wsSql & " AND IPHDUPDFLG = 'N'"
 '   wsSql = wsSql & " AND IPHDSTATUS = '1'"
 '   wsSql = wsSql & " AND IPHDPGMNO = 'AP001' "
    wsSQL = wsSQL & " GROUP BY IPHDDOCNO, IPHDJOBNO, IPHDDOCDATE, IPHDCURR "
 '   wsSql = wsSql & " HAVING SUM(IPDTSTLAMT) = 0 AND SUM(IPDTSTLAMTL) = 0"
    wsSQL = wsSQL & " ORDER BY IPHDDOCNO "
  
    
    Call Ini_Combo(6, wsSQL, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
Dim wsDocDate As String
Dim wsPgmNo As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoGen("PV") = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNo.SetFocus
        Exit Function
    End If
    
        
   If Chk_DocNo(cboDocNo, wsStatus, wsTrnCode, wsUpdFlg, wsDocDate, wsPgmNo) = True Then
        
        If wsTrnCd <> wsTrnCode Then
            gsMsg = "This is not a valid transaction code!"
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
    
        If Chk_ValidDocDate(wsDocDate, "AP") = False Then
            cboDocNo.SetFocus
            Exit Function
        End If
    
    
        If wsPgmNo <> wsFormID Then
            gsMsg = "This is not a valid form code!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
            Chk_cboDocNo = True
            Exit Function
        End If
    
        wsSQL = "SELECT SUM(IPDTSTLAMT) STLAMT FROM APIPHD, APIPDT "
        wsSQL = wsSQL & " WHERE IPHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND IPHDDOCID = IPDTDOCID "
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "STLAMT")) <> 0 Then
            gsMsg = "Document has been already settled!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            rsRcd.Close
            Set rsRcd = Nothing
            wbReadOnly = True
            Chk_cboDocNo = True
            Exit Function
        End If
        End If
        rsRcd.Close
        Set rsRcd = Nothing
    
    
    Else
    
        wsSQL = "SELECT APCQCHQID FROM APCHEQUE "
        wsSQL = wsSQL & " WHERE APCQCHQNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND APCQPGMNO <> '" & wsFormID & "' "
    
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used by cheque!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
        rsRcd.Close
        Set rsRcd = Nothing
        
        wsSQL = "SELECT APSHDOCID FROM APSTHD "
        wsSQL = wsSQL & " WHERE APSHDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND APSHPGMNO <> '" & wsFormID & "' "
    
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used by Settlement!"
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
        txtRevNo.Text = Format(0, "##0")
        txtRevNo.Enabled = False
        medDocDate.Text = Dsp_Date(Now)
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
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
   
        wsSQL = "SELECT IPHDDOCID, IPHDDOCNO, IPHDVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, "
        wsSQL = wsSQL & "IPHDDOCDATE, IPHDREVNO, IPHDCURR, IPHDEXCR, "
        wsSQL = wsSQL & "IPHDDUEDATE, IPHDPAYCODE, IPHDSALEID, IPHDMLCODE, IPHDJOBNO, "
        wsSQL = wsSQL & "IPHDRMKCODE, IPHDRMK1,  IPHDRMK2,  IPHDRMK3,  IPHDRMK4, IPHDRMK5, "
        wsSQL = wsSQL & "IPDTID, IPDTMLCODE, IPDTCATCODE, IPDTCUSPO, IPDTJOBNO, IPDTINVAMT, IPDTDESC "
        wsSQL = wsSQL & "FROM  APIPHD, APIPDT, mstVENDOR "
        wsSQL = wsSQL & "WHERE IPHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND IPHDDOCID = IPDTDOCID "
        wsSQL = wsSQL & "AND IPHDVdrID = VdrID "
    '    wsSql = wsSql & "AND IPHDPGMNO = '" & wsFormID & "'"
        wsSQL = wsSQL & "AND IPHDSTATUS <> '2'"
        wsSQL = wsSQL & "ORDER BY IPDTDOCLINE "
  
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsRcd, "IPHDDOCID")
    txtRevNo.Text = Format(ReadRs(rsRcd, "IPHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsRcd, "IPHDREVNO"))
    medDocDate.Text = ReadRs(rsRcd, "IPHDDOCDATE")
    wlVdrID = ReadRs(rsRcd, "VdrID")
    cboVdrCode.Text = ReadRs(rsRcd, "VdrCode")
    lblDspVdrName.Caption = ReadRs(rsRcd, "VdrName")
    lblDspVdrTel.Caption = ReadRs(rsRcd, "VdrTel")
    lblDspVdrFax.Caption = ReadRs(rsRcd, "VdrFax")
    cboCurr.Text = ReadRs(rsRcd, "IPHDCURR")
    txtExcr.Text = Format(ReadRs(rsRcd, "IPHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsRcd, "IPHDDUEDATE"))
      
     
    cboPayCode = ReadRs(rsRcd, "IPHDPAYCODE")
    cboMLCode = ReadRs(rsRcd, "IPHDMLCODE")
    cboRmkCode = ReadRs(rsRcd, "IPHDRMKCODE")
    cboJobNo = ReadRs(rsRcd, "IPHDJOBNO")
    
    
    
    Dim i As Integer
    
    For i = 1 To 5
        txtRmk(i) = ReadRs(rsRcd, "IPHDRMK" & i)
    Next i
    
    
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, GMLCODE, GDTLID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), GMLCODE) = ReadRs(rsRcd, "IPDTMLCODE")
             waResult(.UpperBound(1), GDESC) = ReadRs(rsRcd, "IPDTDESC")
             waResult(.UpperBound(1), GCUSPO) = ReadRs(rsRcd, "IPDTCUSPO")
             waResult(.UpperBound(1), GJOBNO) = ReadRs(rsRcd, "IPDTJOBNO")
             waResult(.UpperBound(1), GAMT) = IIf(wsTrnCd <> "20", Format(ReadRs(rsRcd, "IPDTINVAMT") * -1, gsAmtFmt), Format(ReadRs(rsRcd, "IPDTINVAMT"), gsAmtFmt))
             waResult(.UpperBound(1), GDTLID) = ReadRs(rsRcd, "IPDTID")
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
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
    lblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblDueDate.Caption = Get_Caption(waScrItm, "DUEDATE")
    lblJobNo.Caption = Get_Caption(waScrItm, "JOBNO")
    
    lblInvAmtOrg.Caption = Get_Caption(waScrItm, "InvAmtORG")
    
    lblInvAmtLoc.Caption = Get_Caption(waScrItm, "InvAmtLOC")
    
    With tblDetail
        .Columns(GMLCODE).Caption = Get_Caption(waScrItm, "GMLCODE")
        .Columns(GDESC).Caption = Get_Caption(waScrItm, "GDESC")
        .Columns(GJOBNO).Caption = Get_Caption(waScrItm, "GJOBNO")
        .Columns(GCUSPO).Caption = Get_Caption(waScrItm, "GCUSPO")
        .Columns(GAMT).Caption = Get_Caption(waScrItm, "GAMT")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    
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
    
    wsActNam(1) = Get_Caption(waScrItm, "APADD")
    wsActNam(2) = Get_Caption(waScrItm, "APEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "APDELETE")
    
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
    Set waScrToolTip = Nothing
    Set waScrItm = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmAP001 = Nothing

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
            cboRmkCode.SetFocus
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
    
        
    If Chk_ValidDocDate(medDueDate.Text, "AP") = False Then
       tabDetailInfo.Tab = 0
        medDueDate.SetFocus
       Exit Function
    End If
    
    Chk_medDueDate = True

End Function





Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If cboPayCode.Enabled Then
            cboPayCode.SetFocus
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
        Select Case wcCombo.Col
          Case GMLCODE
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
              Case GMLCODE
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
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT IPHDSTATUS FROM APIPHD WHERE IPHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
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
        
    adcmdSave.CommandText = "USP_AP001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wsSrcCd)
    Call SetSPPara(adcmdSave, 4, wlKey)
    Call SetSPPara(adcmdSave, 5, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 6, wlVdrID)
    Call SetSPPara(adcmdSave, 7, medDocDate.Text)
    Call SetSPPara(adcmdSave, 8, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 9, cboCurr.Text)
    Call SetSPPara(adcmdSave, 10, txtExcr.Text)
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medDueDate.Text))
    
    Call SetSPPara(adcmdSave, 12, 0)
    
    Call SetSPPara(adcmdSave, 13, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 14, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 15, cboRmkCode.Text)
    Call SetSPPara(adcmdSave, 16, cboJobNo.Text)
    
    
    
    Call SetSPPara(adcmdSave, 17, txtRmk(1).Text)
    Call SetSPPara(adcmdSave, 18, txtRmk(2).Text)
    Call SetSPPara(adcmdSave, 19, txtRmk(3).Text)
    Call SetSPPara(adcmdSave, 20, txtRmk(4).Text)
    Call SetSPPara(adcmdSave, 21, txtRmk(5).Text)
    Call SetSPPara(adcmdSave, 22, "")
    Call SetSPPara(adcmdSave, 23, "")
    Call SetSPPara(adcmdSave, 24, "")
    Call SetSPPara(adcmdSave, 25, "")
    Call SetSPPara(adcmdSave, 26, "")
    
    
    
    Call SetSPPara(adcmdSave, 27, wsFormID)
    Call SetSPPara(adcmdSave, 28, gsUserID)
    Call SetSPPara(adcmdSave, 29, wsGenDte)
    Call SetSPPara(adcmdSave, 30, wsDteTim)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 31)
    wsDocNo = GetSPPara(adcmdSave, 32)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_AP001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GMLCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, GMLCODE))
                Call SetSPPara(adcmdSave, 4, wiCtr + 1)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, GDESC))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, GJOBNO))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, GCUSPO))
                wdTmpAmt = NBRnd(IIf(wsTrnCd <> "20", To_Value(waResult(wiCtr, GAMT)) * -1, To_Value(waResult(wiCtr, GAMT))), giAmtDp)
                Call SetSPPara(adcmdSave, 8, wdTmpAmt)
                wdTmpAmt = NBRnd(IIf(wsTrnCd <> "20", To_Value(waResult(wiCtr, GAMT)) * To_Value(txtExcr.Text) * -1, To_Value(waResult(wiCtr, GAMT)) * To_Value(txtExcr.Text)), giAmtDp)
                Call SetSPPara(adcmdSave, 9, wdTmpAmt)
                Call SetSPPara(adcmdSave, 10, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 11, wsFormID)
                Call SetSPPara(adcmdSave, 12, gsUserID)
                Call SetSPPara(adcmdSave, 13, wsGenDte)
                
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
    
    
    
    If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboVdrCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboMLCode Then Exit Function
    
    If Not Chk_medDueDate Then Exit Function
    
    If Not Chk_cboRmkCode Then Exit Function
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, GMLCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
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

    Dim newForm As New frmAP001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmAP001
    
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
   ' wsFormID = "AP001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsSrcCd = "AP"
   ' wsTrnCd = "20"
    
    wsSOPFlg = Get_SystemFlag("SYPINTSOP")


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

Private Sub txtRevNo_GotFocus()
FocusMe txtRevNo
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
        gsMsg = "修改號錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtRevNo.SetFocus
        Exit Function
    End If
    
    chk_txtRevNo = True

End Function

Private Sub cboVdrCode_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrCode
    
    If gsLangID = "1" Then
        wsSQL = "SELECT VdrCode, VdrName FROM mstVendor "
        wsSQL = wsSQL & "WHERE VdrCode LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSQL = wsSQL & "AND VdrSTATUS = '1' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY VdrCode "
    Else
        wsSQL = "SELECT VdrCode, VdrName FROM mstVendor "
        wsSQL = wsSQL & "WHERE VdrCode LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSQL = wsSQL & "AND VdrSTATUS = '1' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY VdrCode "
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
           
            cboJobNo.SetFocus
            
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
    wsSQL = wsSQL & "FROM  mstVendor "
    wsSQL = wsSQL & "WHERE VdrID = " & wlVdrID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "VDRCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "VDRPAYCODE")
        cboMLCode.Text = ReadRs(rsDefVal, "VDRMLCODE")
        
          Else
        cboCurr.Text = ""
        cboPayCode.Text = ""
        cboMLCode.Text = ""
        
        
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
    
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    
    
    'get Due Date Payment Term
    medDueDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))

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
        
        For wiCtr = GMLCODE To GDummy
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case GMLCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                Case GDESC
                    .Columns(wiCtr).Width = 4300
                    .Columns(wiCtr).DataWidth = 100
                    
                Case GJOBNO
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                    '.Columns(wiCtr).Visible = False
                    
                Case GCUSPO
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).DataWidth = 20
                    
                Case GAMT
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GDTLID
                
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case GDummy
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


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case GMLCODE
                
                If Chk_grdMLCode(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(GJOBNO).Text = cboJobNo.Text
                
             Case GJOBNO
                
                If Chk_grdJobNo(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case GAMT
                                
                If Chk_Amount(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(ColIndex).Text = Format(NBRnd(.Columns(ColIndex).Text, giAmtDp), gsAmtFmt)
                
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
            Case GMLCODE
                
                wsSQL = "SELECT MLCODE, MLDESC FROM mstMerchClass "
                wsSQL = wsSQL & " WHERE MLSTATUS <> '2' "
                wsSQL = wsSQL & " AND MLCODE LIKE '%" & Set_Quote(.Columns(GMLCODE).Text) & "%' "
                wsSQL = wsSQL & " AND MLTYPE = 'P' "
                wsSQL = wsSQL & " ORDER BY MLCODE "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case GJOBNO
                
            If wsSOPFlg = "Y" Then
                
                wsSQL = "SELECT SOHDDOCNO, CUSCODE FROM SOASOHD, MSTCUSTOMER "
                wsSQL = wsSQL & " WHERE SOHDSTATUS IN ('1','4') "
                wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(GJOBNO).Text) & "%' "
                wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
                wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
            Else
                wsSQL = "SELECT JOBCODE, JOBNAME FROM mstJOB "
                wsSQL = wsSQL & " WHERE JOBSTATUS <> '2' AND JOBCODE LIKE '%" & Set_Quote(.Columns(GJOBNO).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY JOBCODE "
                
            End If
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLJOBCODE", Me.Width, Me.Height)
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
                Case GMLCODE, GCUSPO, GJOBNO
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case GDESC
                KeyCode = vbDefault
                    .Col = GJOBNO
                Case GAMT
                    KeyCode = vbKeyDown
                    .Col = GMLCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> GMLCODE Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> GAMT Then
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
        
        Case GAMT
            Call Chk_InpNum(KeyAscii, tblDetail.Text, True, True)
        
      
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = GMLCODE
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case GMLCODE
                    Call Chk_grdMLCode(.Columns(GMLCODE))
                Case GJOBNO
                    Call Chk_grdJobNo(.Columns(GJOBNO).Text)
                Case GAMT
                    Call Chk_Amount(.Columns(GAMT).Text)
                
            End Select
        End If
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
    wsSQL = wsSQL & " AND MLTYPE = 'P' "
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



Private Function Chk_grdJobNo(inNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdJobNo = False
    
  '  If Trim(inNo) = "" Then
        Chk_grdJobNo = True
        Exit Function
  '  End If
    
    If wsSOPFlg = "Y" Then
    
    wsSQL = "SELECT * FROM SOASOHD "
    wsSQL = wsSQL & " WHERE SOHDDOCNO = '" & Set_Quote(inNo) & "' "
    wsSQL = wsSQL & " AND SOHDSTATUS = '4' "
    
    Else
    
    wsSQL = "SELECT *  FROM mstJob "
    wsSQL = wsSQL & " WHERE JobCode = '" & Set_Quote(inNo) & "' "
    wsSQL = wsSQL & " AND JOBSTATUS = '1' "
    
    End If
       
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        gsMsg = "沒有此工程!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_grdJobNo = True
        

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
                If Trim(.Columns(GMLCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, GMLCODE)) = "" And _
                   Trim(waResult(inRow, GDESC)) = "" And _
                   Trim(waResult(inRow, GJOBNO)) = "" And _
                   Trim(waResult(inRow, GCUSPO)) = "" And _
                   Trim(waResult(inRow, GAMT)) = "" And _
                   Trim(waResult(inRow, GDTLID)) = "" Then
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
        
        If Chk_grdMLCode(waResult(LastRow, GMLCODE)) = False Then
            .Col = GMLCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdJobNo(waResult(LastRow, GJOBNO)) = False Then
                .Col = GJOBNO
                .Row = LastRow
                Exit Function
        End If
        
        
        If Chk_Amount(waResult(LastRow, GAMT)) = False Then
            .Col = GAMT
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
    
    Dim wiTotal As Double
    
    Dim wiRowCtr As Integer
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotal = wiTotal + To_Value(waResult(wiRowCtr, GAMT))
    Next
    
    lblDspInvAmtOrg.Caption = Format(CStr(wiTotal), gsAmtFmt)
    lblDspInvAmtLoc.Caption = Format(CStr(wiTotal * To_Value(txtExcr)), gsAmtFmt)
    
    Calc_Total = True

End Function




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
    
    gsMsg = "你是否確認要刪除此檔案?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       wiAction = CorRec
       MousePointer = vbDefault
       Exit Function
    End If
    
    wiAction = DelRec
    
      cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_AP001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
     Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wsSrcCd)
    Call SetSPPara(adcmdDelete, 4, wlKey)
    Call SetSPPara(adcmdDelete, 5, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 6, wlVdrID)
    Call SetSPPara(adcmdDelete, 7, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 8, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 9, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 10, txtExcr.Text)
    
    Call SetSPPara(adcmdDelete, 11, Set_MedDate(medDueDate.Text))
    
    Call SetSPPara(adcmdDelete, 12, 0)
    
    Call SetSPPara(adcmdDelete, 13, cboMLCode.Text)
    Call SetSPPara(adcmdDelete, 14, cboPayCode.Text)
    Call SetSPPara(adcmdDelete, 15, cboRmkCode.Text)
    Call SetSPPara(adcmdDelete, 16, cboJobNo.Text)
    
    
    Call SetSPPara(adcmdDelete, 17, txtRmk(1).Text)
    Call SetSPPara(adcmdDelete, 18, txtRmk(2).Text)
    Call SetSPPara(adcmdDelete, 19, txtRmk(3).Text)
    Call SetSPPara(adcmdDelete, 20, txtRmk(4).Text)
    Call SetSPPara(adcmdDelete, 21, txtRmk(5).Text)
    Call SetSPPara(adcmdDelete, 22, "")
    Call SetSPPara(adcmdDelete, 23, "")
    Call SetSPPara(adcmdDelete, 24, "")
    Call SetSPPara(adcmdDelete, 25, "")
    Call SetSPPara(adcmdDelete, 26, "")
    
    
    
    Call SetSPPara(adcmdDelete, 27, wsFormID)
    Call SetSPPara(adcmdDelete, 28, gsUserID)
    Call SetSPPara(adcmdDelete, 29, wsGenDte)
    Call SetSPPara(adcmdDelete, 30, wsDteTim)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 31)
    wsDocNo = GetSPPara(adcmdDelete, 32)
   
    
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
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medDueDate.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.cboRmkCode.Enabled = False
            Me.cboJobNo.Enabled = False
            
            
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
            Me.cboPayCode.Enabled = True
            Me.cboMLCode.Enabled = True
            Me.cboRmkCode.Enabled = True
            Me.cboJobNo.Enabled = True
            
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
        .TableType = wsSrcCd
        .TableKey = "INHDDocNo"
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
    vFilterAry(1, 2) = "IPHDDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "IPHDDocDate"
    
    vFilterAry(3, 1) = "Vendor #"
    vFilterAry(3, 2) = "VdrCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "IPHDDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "IPHDDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Vendor#"
    vAry(3, 2) = "VdrCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Vendor Name"
    vAry(4, 2) = "VdrName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT IpHdDocNo, IphdDocDate, mstVendor.VdrCode,  mstVendor.VdrName "
        wsSQL = wsSQL + "FROM MstVendor, APIPHD "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE IpHdStatus = '1' And IpHdVdrID = VdrID "
        .sBindOrderSQL = "ORDER BY IpHdDocNo"
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
        cboMLCode.SetFocus
       
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
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass "
    wsSQL = wsSQL & " WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSQL = wsSQL & " AND MLSTATUS = '1' "
    wsSQL = wsSQL & " AND MLTYPE = 'R' "
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
    
    
    If Chk_MLClass(cboMLCode, "R", wsDesc) = False Then
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





Private Sub txtRevNo_LostFocus()
    FocusMe txtRevNo, True
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
        
        tabDetailInfo.Tab = 0
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
        tabDetailInfo.Tab = 0
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
        
        
        If Index = 5 Then
        tabDetailInfo.Tab = 1
        tblDetail.SetFocus
        Else
        tabDetailInfo.Tab = 0
        txtRmk(Index + 1).SetFocus
        End If
        
    End If
End Sub

Private Sub txtRmk_LostFocus(Index As Integer)
        
        FocusMe txtRmk(Index), True

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
        
        For i = 1 To 5
        txtRmk(i) = ReadRs(rsRcd, "RmkDESC" & i)
        Next i
        
        Else
        
        For i = 1 To 5
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
Private Function Chk_DocNo(ByVal InDocNo As String, ByRef OutStatus As String, ByRef OutTrnCd As String, ByRef OutUpdFlg As String, ByRef OutDocDate As String, ByRef OutPgmNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    OutStatus = ""
    OutTrnCd = ""
    OutUpdFlg = ""
    OutDocDate = ""
    Chk_DocNo = False
    
    wsSQL = "SELECT IPHDTRNCODE, IPHDSTATUS, IPHDUPDFLG, IPHDDOCDATE, IPHDPGMNO FROM APIPHD "
    wsSQL = wsSQL & " WHERE IPHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount <= 0 Then
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    OutStatus = ReadRs(rsRcd, "IPHDSTATUS")
    OutTrnCd = ReadRs(rsRcd, "IPHDTRNCODE")
    OutUpdFlg = ReadRs(rsRcd, "IPHDUPDFLG")
    OutDocDate = ReadRs(rsRcd, "IPHDDOCDATE")
    OutPgmNo = ReadRs(rsRcd, "IPHDPGMNO")
    
    rsRcd.Close
    Set rsRcd = Nothing
    
       
   
    Chk_DocNo = True
   

End Function


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property



Private Sub cboJobNo_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboJobNo
    
    wsSQL = "SELECT SOHDDOCNO, SOHDSHIPFROM FROM SOASOHD "
    wsSQL = wsSQL & "WHERE SOHDDOCNO LIKE '%" & IIf(cboJobNo.SelLength > 0, "", Set_Quote(cboJobNo.Text)) & "%' "
    wsSQL = wsSQL & "AND SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & "ORDER BY SOHDDOCNO "
    
    Call Ini_Combo(2, wsSQL, cboJobNo.Left, cboJobNo.Top + cboJobNo.Height, tblCommon, wsFormID, "TBLJOBCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
   
End Sub

Private Sub cboJobNo_GotFocus()
    
    Set wcCombo = cboJobNo
    FocusMe cboJobNo
    
End Sub

Private Sub cboJobNo_LostFocus()
    FocusMe cboJobNo, True
End Sub

Private Sub cboJobNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboJobNo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_cboJobNo() = False Then Exit Sub
           
            cboCurr.SetFocus
            
    End If
    
End Sub

Private Function Chk_cboJobNo() As Boolean
    
    
    Chk_cboJobNo = False
        
    If Trim(cboJobNo) = "" Then
        Chk_cboJobNo = True
        Exit Function
    End If
    

    
    Chk_cboJobNo = True

End Function


