VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmWS001 
   BackColor       =   &H8000000A&
   Caption         =   "WSINFO"
   ClientHeight    =   5670
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "frmWS001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "frmWS001.frx":08CA
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboWsSaleCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox cboWsVdrCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   31
      Top             =   3480
      Width           =   1335
   End
   Begin VB.ComboBox cboWsCusCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   30
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox cboWSWhsCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   23
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cboWSMethodCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   20
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cboWSPayCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   17
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cboCurr 
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cboWSID 
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtWSExcr 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtWSID 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Tag             =   "K"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblDspWsSaleName 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   29
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Label lblWSSale 
         Caption         =   "WSSALE"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   3540
         Width           =   2100
      End
      Begin VB.Label lblDspWsVdrName 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   27
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Label lblWSVdr 
         Caption         =   "WSVDR"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   3180
         Width           =   2100
      End
      Begin VB.Label lblDspWsCusName 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   25
         Top             =   2760
         Width           =   4335
      End
      Begin VB.Label lblWSCus 
         Caption         =   "WSCUS"
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   2820
         Width           =   2100
      End
      Begin VB.Label lblDspWSWhsDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   22
         Top             =   2400
         Width           =   4335
      End
      Begin VB.Label lblWSWhsCode 
         Caption         =   "WHSCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   2460
         Width           =   2100
      End
      Begin VB.Label lblDspWSMethodDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   19
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label lblWSMethodCode 
         Caption         =   "METHODCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   2100
         Width           =   2100
      End
      Begin VB.Label lblDspWSPayDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3840
         TabIndex        =   16
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label lblWSPayCode 
         Caption         =   "PAYCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1740
         Width           =   2100
      End
      Begin VB.Label lblWSExcr 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   2100
      End
      Begin VB.Label lblCurr 
         Caption         =   "CURR"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   1035
         Width           =   2100
      End
      Begin VB.Label lblWSLastUpd 
         Caption         =   "WSLASTUPD"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   4725
         Width           =   2340
      End
      Begin VB.Label lblWSLastUpdDate 
         Caption         =   "WSLASTUPDDATE"
         Height          =   240
         Left            =   4080
         TabIndex        =   8
         Top             =   4725
         Width           =   2460
      End
      Begin VB.Label lblDspWSLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   4680
         Width           =   1305
      End
      Begin VB.Label lblDspWSLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6840
         TabIndex        =   6
         Top             =   4680
         Width           =   1305
      End
      Begin VB.Label lblWSID 
         Caption         =   "WSID"
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
         TabIndex        =   5
         Top             =   660
         Width           =   2100
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
            Picture         =   "frmWS001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWS001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
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
Attribute VB_Name = "frmWS001"
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
Private wlMLAccID As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private wlCusID As Long
Private wlVdrID As Long
Private wlSaleID As Long

Private Const wsKeyType = "sysWSINFO"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboWsCusCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWsCusCode
    
    If gsLangID = "1" Then
        wsSql = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSql = wsSql & "WHERE CUSCODE LIKE '%" & IIf(cboWsCusCode.SelLength > 0, "", Set_Quote(cboWsCusCode.Text)) & "%' "
        wsSql = wsSql & "AND CUSSTATUS = '1' "
        wsSql = wsSql & " AND CusInactive = 'N' "
        wsSql = wsSql & "ORDER BY CUSCODE "
    Else
        wsSql = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSql = wsSql & "WHERE CUSCODE LIKE '%" & IIf(cboWsCusCode.SelLength > 0, "", Set_Quote(cboWsCusCode.Text)) & "%' "
        wsSql = wsSql & "AND CUSSTATUS = '1' "
        wsSql = wsSql & " AND CusInactive = 'N' "
        wsSql = wsSql & "ORDER BY CUSCODE "
    End If
    Call Ini_Combo(2, wsSql, cboWsCusCode.Left, cboWsCusCode.Top + cboWsCusCode.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWsCusCode_GotFocus()
    FocusMe cboWsCusCode
End Sub

Private Sub cboWsCusCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWsCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_cboWsCusCode() Then
            cboWsVdrCode.SetFocus
        End If
    End If
End Sub

Private Sub cboWsCusCode_LostFocus()
    FocusMe cboWsCusCode, True
End Sub

Private Sub cboWSID_LostFocus()
    FocusMe cboWSID, True
End Sub

Private Sub cboWSMethodCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboWSMethodCode
    
    wsSql = "SELECT MethodCode, MethodDesc FROM MstMethod WHERE MethodStatus = '1'"
    wsSql = wsSql & " AND MethodCode LIKE '%" & IIf(cboWSMethodCode.SelLength > 0, "", Set_Quote(cboWSMethodCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY MethodCode "
    Call Ini_Combo(2, wsSql, cboWSMethodCode.Left, cboWSMethodCode.Top + cboWSMethodCode.Height, tblCommon, wsFormID, "TBLM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWSMethodCode_GotFocus()
    FocusMe cboWSMethodCode
End Sub

Private Sub cboWSMethodCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWSMethodCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboWSMethodCode() = True Then
            cboWSWhsCode.SetFocus
        End If
    End If
End Sub

Private Sub cboWSMethodCode_LostFocus()
    FocusMe cboWSMethodCode, True
End Sub

Private Sub cboWSPayCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboWSPayCode
    
    wsSql = "SELECT PayCode, PayDesc, PayMethod FROM MstPayTerm WHERE PayStatus = '1'"
    wsSql = wsSql & " AND PayCode LIKE '%" & IIf(cboWSPayCode.SelLength > 0, "", Set_Quote(cboWSPayCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY PayCode "
    Call Ini_Combo(3, wsSql, cboWSPayCode.Left, cboWSPayCode.Top + cboWSPayCode.Height, tblCommon, wsFormID, "TBLPAY", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWSPayCode_GotFocus()
    FocusMe cboWSPayCode
End Sub

Private Sub cboWSPayCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWSPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboWSPayCode() = True Then
            cboWSMethodCode.SetFocus
        End If
    End If
End Sub

Private Sub cboWSPayCode_LostFocus()
    FocusMe cboWSPayCode, True
End Sub

Private Sub cboWsSaleCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboWsSaleCode
    
    wsSql = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleStatus = '1'"
    wsSql = wsSql & " AND SaleCode LIKE '%" & IIf(cboWsSaleCode.SelLength > 0, "", Set_Quote(cboWsSaleCode.Text)) & "%' "
   
    wsSql = wsSql & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSql, cboWsSaleCode.Left, cboWsSaleCode.Top + cboWsSaleCode.Height, tblCommon, wsFormID, "TBLSLMCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWsSaleCode_GotFocus()
    FocusMe cboWsSaleCode
End Sub

Private Sub cboWsSaleCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWsSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_cboWSSaleCode() Then
            cboCurr.SetFocus
        End If
    End If
End Sub

Private Sub cboWsSaleCode_LostFocus()
    FocusMe cboWsSaleCode, True
End Sub

Private Sub cboWsVdrCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWsVdrCode
    
    If gsLangID = "1" Then
        wsSql = "SELECT VDRCODE, VDRNAME FROM MstVendor "
        wsSql = wsSql & "WHERE VDRCODE LIKE '%" & IIf(cboWsVdrCode.SelLength > 0, "", Set_Quote(cboWsVdrCode.Text)) & "%' "
        wsSql = wsSql & "AND VDrSTATUS = '1' "
        wsSql = wsSql & " AND VdrInactive = 'N' "
        wsSql = wsSql & "ORDER BY VDRCODE "
    Else
        wsSql = "SELECT VDRCODE, VDRNAME FROM MstVendor "
        wsSql = wsSql & "WHERE VDRCODE LIKE '%" & IIf(cboWsVdrCode.SelLength > 0, "", Set_Quote(cboWsVdrCode.Text)) & "%' "
        wsSql = wsSql & "AND VDRSTATUS = '1' "
        wsSql = wsSql & " AND VdrInactive = 'N' "
        wsSql = wsSql & "ORDER BY VDrCODE "
    End If
    Call Ini_Combo(2, wsSql, cboWsVdrCode.Left, cboWsVdrCode.Top + cboWsVdrCode.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWsVdrCode_GotFocus()
    FocusMe cboWsVdrCode
End Sub

Private Sub cboWsVdrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWsVdrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_cboWsVdrCode() Then
            cboWsSaleCode.SetFocus
        End If
    End If
End Sub

Private Sub cboWsVdrCode_LostFocus()
    FocusMe cboWsVdrCode, True
End Sub

Private Sub cboWSWhsCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboWSWhsCode
    
    wsSql = "SELECT WhsCode, WhsDesc FROM MstWarehouse WHERE WhsStatus = '1'"
    wsSql = wsSql & " AND WhsCode LIKE '%" & IIf(cboWSWhsCode.SelLength > 0, "", Set_Quote(cboWSWhsCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY WhsCode "
    Call Ini_Combo(2, wsSql, cboWSWhsCode.Left, cboWSWhsCode.Top + cboWSWhsCode.Height, tblCommon, wsFormID, "TBLWHS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWSWhsCode_GotFocus()
    FocusMe cboWSWhsCode
End Sub

Private Sub cboWSWhsCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWSWhsCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboWSWhsCode() = True Then
            cboWsCusCode.SetFocus
        End If
    End If
End Sub

Private Sub cboWSWhsCode_LostFocus()
    FocusMe cboWSWhsCode, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
        Me.Height = 6075
        Me.Width = 8700
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
            Me.cboWSID.Enabled = False
            Me.cboWSID.Visible = False
            Me.txtWSID.Visible = True
            Me.txtWSID.Enabled = False
            
            cboCurr.Enabled = False
            txtWSExcr.Enabled = False
            cboWSPayCode.Enabled = False
            cboWSMethodCode.Enabled = False
            cboWSWhsCode.Enabled = False
            cboWsCusCode.Enabled = False
            cboWsVdrCode.Enabled = False
            cboWsSaleCode.Enabled = False
            
        Case "AfrActAdd"
            Me.cboWSID.Enabled = False
            Me.cboWSID.Visible = False
            
            Me.txtWSID.Enabled = True
            Me.txtWSID.Visible = True
            
        Case "AfrActEdit"
            Me.cboWSID.Enabled = True
            Me.cboWSID.Visible = True
            
            Me.txtWSID.Enabled = False
            Me.txtWSID.Visible = False
            
        Case "AfrKey"
            Me.cboWSID.Enabled = False
            Me.txtWSID.Enabled = False
            
            cboCurr.Enabled = True
            txtWSExcr.Enabled = True
            cboWSPayCode.Enabled = True
            cboWSMethodCode.Enabled = True
            cboWSWhsCode.Enabled = True
            cboWsCusCode.Enabled = True
            cboWsVdrCode.Enabled = True
            cboWsSaleCode.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_cboCurr = False Then
        Exit Function
    End If
    
    If Chk_txtWSExcr = False Then
        Exit Function
    End If
    
    If Chk_cboWSPayCode = False Then
        Exit Function
    End If
    
    If Chk_cboWSMethodCode = False Then
        Exit Function
    End If
    
    If Chk_cboWSWhsCode = False Then
        Exit Function
    End If
    
    If chk_cboWsCusCode = False Then
        Exit Function
    End If
    
    If chk_cboWsVdrCode = False Then
        Exit Function
    End If
    
    If Chk_cboWSSaleCode = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT sysWSINFO.* "
    wsSql = wsSql + "From sysWSINFO "
    wsSql = wsSql + "WHERE (((sysWSINFO.WSID)='" + Set_Quote(cboWSID) + "') "
    wsSql = wsSql + "AND ((WSStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        
        cboWSID = ReadRs(rsRcd, "WSID")
        cboCurr = ReadRs(rsRcd, "WSCURR")
        txtWSExcr = Format(ReadRs(rsRcd, "WSEXCR"), gsExrFmt)
        cboWSPayCode = ReadRs(rsRcd, "WSPAYCODE")
        cboWSMethodCode = ReadRs(rsRcd, "WSMETHODCODE")
        cboWSWhsCode = ReadRs(rsRcd, "WSWHSCODE")
        wlCusID = ReadRs(rsRcd, "WSCUSID")
        wlVdrID = ReadRs(rsRcd, "WSVDRID")
        wlSaleID = ReadRs(rsRcd, "WSSALEID")
        
        cboWsSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
        lblDspWsSaleName = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
        cboWsVdrCode.Text = Get_TableInfo("MstVendor", "VDRID =" & wlVdrID, "VDRCODE")
        lblDspWsVdrName = Get_TableInfo("MstVendor", "VDRID =" & wlVdrID, "VDRNAME")
        cboWsCusCode.Text = Get_TableInfo("MstCustomer", "CUSID =" & wlCusID, "CUSCODE")
        lblDspWsCusName = Get_TableInfo("MstCustomer", "CUSID =" & wlCusID, "CUSNAME")
        lblDspWSPayDesc = Get_TableInfo("MstPayTerm", "PayCode ='" & cboWSPayCode & "'", "PAYDESC")
        lblDspWSMethodDesc = Get_TableInfo("MstMethod", "MethodCode ='" & cboWSMethodCode & "'", "METHODDESC")
        lblDspWSWhsDesc = Get_TableInfo("MstWareHouse", "WhsCode ='" & cboWSWhsCode & "'", "WHSDESC")
        
        lblDspWSLastUpd = ReadRs(rsRcd, "WSLastUpd")
        lblDspWSLastUpdDate = ReadRs(rsRcd, "WSLastUpdDate")
        
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
    Call UnlockAll(wsConnTime, wsFormID)
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set frmWS001 = Nothing
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
'    Me.Left = 0
'    Me.Top = 0
'    Me.Width = Screen.Width
'    Me.Height = Screen.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "WS001"
    wsTrnCd = ""
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
    
    txtWSExcr = Format(0, gsExrFmt)
    
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
        txtWSID.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboWSID.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboWSID.SetFocus
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
            If RowLock(wsConnTime, wsKeyType, cboWSID, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    cboCurr.SetFocus
End Sub

Private Function Chk_WSID(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_WSID = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT WSStatus "
    wsSql = wsSql & " FROM sysWSINFO WHERE WSID = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "WSStatus")
    
    Chk_WSID = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_Curr(ByVal inCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT ExcCurr FROM MstExchangeRate WHERE ExcCurr='" & Set_Quote(inCode) & "' And ExcStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount < 1 Then
        Chk_Curr = False
    Else
        Chk_Curr = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboWSID() As Boolean
    Dim wsStatus As String

    Chk_cboWSID = False
    
    If Trim(cboWSID.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWSID.SetFocus
        Exit Function
    End If

    If Chk_WSID(cboWSID.Text, wsStatus) = False Then
        gsMsg = "工作站編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWSID.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "工作站已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboWSID.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboWSID = True
End Function

Private Function Chk_cboCurr() As Boolean
    Dim wsStatus As String

    Chk_cboCurr = False
    
    If Trim(cboCurr.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCurr.SetFocus
        Exit Function
    End If

    If Chk_Curr(cboCurr.Text) = False Then
        gsMsg = "貨幣編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCurr.SetFocus
        Exit Function
    End If
    
    Chk_cboCurr = True
End Function

Private Function Chk_cboWSPayCode() As Boolean
    Dim wsDesc As String

    Chk_cboWSPayCode = False
    
    If Chk_PayTerm(cboWSPayCode.Text, wsDesc) = False Then
        gsMsg = "付款條款不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWSPayCode.SetFocus
        Exit Function
    End If
    
    Me.lblDspWSPayDesc = wsDesc
    
    Chk_cboWSPayCode = True
End Function

Private Function Chk_txtWSID() As Boolean
    Dim wsStatus As String
    
    Chk_txtWSID = False
    
    If Trim(txtWSID.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtWSID.SetFocus
        Exit Function
    End If
    
    If Chk_WSID(txtWSID.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "工作站編碼已存在但已無效!"
        Else
            gsMsg = "工作站編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtWSID.SetFocus
        Exit Function
    End If
    
    Chk_txtWSID = True
End Function

Private Function Chk_txtWSExcr() As Boolean
    Dim wsStatus As String
    
    Chk_txtWSExcr = False
    
    If To_Value(txtWSExcr) < 0 Then
        gsMsg = "對換率不可少於零!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtWSExcr.SetFocus
        Exit Function
    End If
    
    If To_Value(txtWSExcr) > 100 Then
        gsMsg = "對換率不可大於一百!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtWSExcr.SetFocus
        Exit Function
    End If
    
    Chk_txtWSExcr = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmWS001
    
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
                Call UnlockAll(wsConnTime, wsFormID)
                Call Ini_Scr
                Call cmdEdit
                
            Case DelRec
                Call UnlockAll(wsConnTime, wsFormID)
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboWSID, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_WS001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtWSID, cboWSID))
    Call SetSPPara(adcmdSave, 3, cboCurr)
    Call SetSPPara(adcmdSave, 4, txtWSExcr)
    Call SetSPPara(adcmdSave, 5, cboWSPayCode)
    Call SetSPPara(adcmdSave, 6, cboWSMethodCode)
    Call SetSPPara(adcmdSave, 7, cboWSWhsCode)
    Call SetSPPara(adcmdSave, 8, wlCusID)
    Call SetSPPara(adcmdSave, 9, wlVdrID)
    Call SetSPPara(adcmdSave, 10, wlSaleID)
    Call SetSPPara(adcmdSave, 11, gsUserID)
    Call SetSPPara(adcmdSave, 12, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 13)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - WS001!"
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
    MsgBox Err.description
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

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim sSQL As String
    
    ReDim vFilterAry(6, 2)
    vFilterAry(1, 1) = "工作站編碼"
    vFilterAry(1, 2) = "WSID"
    
    vFilterAry(2, 1) = "貨幣"
    vFilterAry(2, 2) = "WSCURR"
    
    vFilterAry(3, 1) = "對換率"
    vFilterAry(3, 2) = "WSEXCR"
    
    vFilterAry(4, 1) = "付款條款"
    vFilterAry(4, 2) = "WSPAYCODE"
    
    vFilterAry(5, 1) = "銷售渠道"
    vFilterAry(5, 2) = "WSMETHODCODE"
    
    vFilterAry(6, 1) = "貨倉"
    vFilterAry(6, 2) = "WSWHSCODE"
    
    ReDim vAry(6, 3)
    vAry(1, 1) = "工作站編碼"
    vAry(1, 2) = "WSID"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "貨幣"
    vAry(2, 2) = "WSCURR"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "對換率"
    vAry(3, 2) = "WSEXCR"
    vAry(3, 3) = "1500"
    
    vAry(4, 1) = "付款條款"
    vAry(4, 2) = "WSPAYCODE"
    vAry(4, 3) = "1500"
    
    vAry(5, 1) = "銷售渠道"
    vAry(5, 2) = "WSMETHODCODe"
    vAry(5, 3) = "1500"
    
    vAry(6, 1) = "貨倉"
    vAry(6, 2) = "WSWHSCODE"
    vAry(6, 3) = "1500"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT sysWSINFO.WSID, sysWSINFO.WSCURR, sysWSINFO.WSEXCR, sysWSINFO.WSPAYCODE, sysWSINFO.WSMETHODCODE, "
        sSQL = sSQL + "sysWSINFO.WSWHSCODE "
        sSQL = sSQL + "FROM sysWSINFO "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE sysWSINFO.WSStatus = '1' "
        .sBindOrderSQL = "ORDER BY sysWSINFO.WSID"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboWSID Then
        cboWSID = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtWSExcr_GotFocus()
    FocusMe txtWSExcr
End Sub

Private Sub txtWSExcr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtWSID, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtWSExcr() = True Then
            cboWSPayCode.SetFocus
        End If
    End If
End Sub

Private Sub txtWSExcr_LostFocus()
    txtWSExcr = Format(To_Value(txtWSExcr), gsExrFmt)
    FocusMe txtWSExcr, True
End Sub

Private Sub txtWSID_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtWSID, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtWSID() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtWSID_LostFocus()
    FocusMe txtWSID, True
End Sub

Private Sub txtWSID_GotFocus()
    FocusMe txtWSID
End Sub

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

Private Sub cboWSID_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWSID, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboWSID() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboWSID_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboWSID
    
    wsSql = "SELECT WSID FROM sysWSINFO WHERE WSStatus = '1'"
    wsSql = wsSql & " AND WSID LIKE '%" & IIf(cboWSID.SelLength > 0, "", Set_Quote(cboWSID.Text)) & "%' "
    wsSql = wsSql & "ORDER BY WSID "
    Call Ini_Combo(1, wsSql, cboWSID.Left, cboWSID.Top + cboWSID.Height, tblCommon, wsFormID, "TBLWS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWSID_GotFocus()
    FocusMe cboWSID
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT WsStatus FROM sysWSINFO WHERE WSID = '" & Set_Quote(txtWSID) & "'"
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
    
    'Create Selection Criteria
    With Newfrm
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "WSID"
        .KeyLen = 10
        Set .ctlKey = txtWSID
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
    
    lblWSID.Caption = Get_Caption(waScrItm, "WSID")
    lblCurr.Caption = Get_Caption(waScrItm, "Curr")
    lblWSExcr.Caption = Get_Caption(waScrItm, "WSEXCR")
    lblWSPayCode.Caption = Get_Caption(waScrItm, "WSPAYCODE")
    lblWSMethodCode.Caption = Get_Caption(waScrItm, "WSMETHODCODE")
    lblWSWhsCode.Caption = Get_Caption(waScrItm, "WSWHSCODE")
    lblWSCus.Caption = Get_Caption(waScrItm, "WSCUS")
    lblWSVdr.Caption = Get_Caption(waScrItm, "WSVDR")
    lblWSSale.Caption = Get_Caption(waScrItm, "WSSALE")
    
    lblWSLastUpd.Caption = Get_Caption(waScrItm, "WSLASTUPD")
    lblWSLastUpdDate.Caption = Get_Caption(waScrItm, "WSLASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "WSADD")
    wsActNam(2) = Get_Caption(waScrItm, "WSEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "WSDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub cboCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCurr() = True Then
            txtWSExcr.SetFocus
        End If
    End If
End Sub

Private Sub cboCurr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCurr
    
    wsSql = "SELECT ExcCurr, ExcDesc FROM MstExchangeRate WHERE ExcStatus = '1'"
    wsSql = wsSql & " AND ExcCurr LIKE '%" & IIf(cboCurr.SelLength > 0, "", Set_Quote(cboCurr.Text)) & "%' "
    wsSql = wsSql & " GROUP BY ExcCurr, ExcDesc "
    wsSql = wsSql & " ORDER BY ExcCurr "
    Call Ini_Combo(2, wsSql, cboCurr.Left, cboCurr.Top + cboCurr.Height, tblCommon, wsFormID, "TBLCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_LostFocus()
    FocusMe cboCurr, True
End Sub

Private Function Chk_cboWSMethodCode() As Boolean
    Dim wsDesc As String

    Chk_cboWSMethodCode = False
    
    If Chk_Method(cboWSMethodCode.Text, wsDesc) = False Then
        gsMsg = "銷售渠道不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWSMethodCode.SetFocus
        Exit Function
    End If
    
    Me.lblDspWSMethodDesc = wsDesc
    
    Chk_cboWSMethodCode = True
End Function

Private Function Chk_cboWSWhsCode() As Boolean
    Dim wsDesc As String

    Chk_cboWSWhsCode = False
    
    If Chk_Whs(cboWSWhsCode.Text, wsDesc) = False Then
        gsMsg = "貨倉不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWSWhsCode.SetFocus
        Exit Function
    End If
    
    lblDspWSWhsDesc = wsDesc
    
    Chk_cboWSWhsCode = True
End Function

Private Function chk_cboWsCusCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    
    chk_cboWsCusCode = False
    
    If Trim(cboWsCusCode) = "" Then
        gsMsg = "必需輸入客戶編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsCusCode.SetFocus
        Exit Function
    End If
    
    If Chk_CusCode(cboWsCusCode, wlID, wsName, "", "") Then
        wlCusID = wlID
        lblDspWsCusName.Caption = wsName
    Else
        wlCusID = 0
        lblDspWsCusName.Caption = ""
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsCusCode.SetFocus
        Exit Function
    End If
    
    chk_cboWsCusCode = True

End Function

Private Function chk_cboWsVdrCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    
    chk_cboWsVdrCode = False
    
    If Trim(cboWsVdrCode) = "" Then
        gsMsg = "必需輸入供應商編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsVdrCode.SetFocus
        Exit Function
    End If
    
    If Chk_VDRCODE(cboWsVdrCode, wlID, wsName, "", "") Then
        wlVdrID = wlID
        lblDspWsVdrName.Caption = wsName
    Else
        wlVdrID = 0
        lblDspWsVdrName.Caption = ""
        gsMsg = "供應商不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsVdrCode.SetFocus
        Exit Function
    End If
    
    chk_cboWsVdrCode = True

End Function

Private Function Chk_cboWSSaleCode() As Boolean
    Dim wsDesc As String

    Chk_cboWSSaleCode = False
     
    If Trim(cboWsSaleCode.Text) = "" Then
        gsMsg = "必需輸入營業員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsSaleCode.SetFocus
        Exit Function
    End If
    
    If Chk_Salesman(cboWsSaleCode, wlSaleID, wsDesc) = False Then
        gsMsg = "沒有此營業員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWsSaleCode.SetFocus
        lblDspWsSaleName = ""
       Exit Function
    End If
    
    lblDspWsSaleName = wsDesc
    
    Chk_cboWSSaleCode = True
    
End Function

