VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPjt001 
   Caption         =   "Project Maintenance"
   ClientHeight    =   7920
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   11460
   Icon            =   "frmPjt001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   11460
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11640
      OleObjectBlob   =   "frmPjt001.frx":030A
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboPjtUOMCode 
      Height          =   300
      Left            =   6960
      TabIndex        =   4
      Top             =   1440
      Width           =   1600
   End
   Begin VB.ComboBox cboPjtTypeCode 
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   1600
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5055
      Left            =   240
      OleObjectBlob   =   "frmPjt001.frx":2A0D
      TabIndex        =   6
      Top             =   2160
      Width           =   11055
   End
   Begin VB.ComboBox cboPjtCurr 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1600
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame fra1 
      Height          =   7335
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtPjtUPrice 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1320
         Width           =   1600
      End
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   5040
         TabIndex        =   17
         Top             =   120
         Width           =   6135
         Begin VB.Label lblKeyDesc 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   360
            TabIndex        =   21
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblComboPrompt 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   1920
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblInsertLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   3360
            TabIndex        =   19
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblDeleteLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   4800
            TabIndex        =   18
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.TextBox txtPjtName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label lblDspPjtUPriceL 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3960
         TabIndex        =   30
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblPjtUOMCode 
         Caption         =   "PJTUOMCODE"
         Height          =   255
         Left            =   5760
         TabIndex        =   29
         Top             =   1035
         Width           =   1200
      End
      Begin VB.Label lblPjtType 
         Caption         =   "PJTYPE"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   1035
         Width           =   1200
      End
      Begin VB.Label lblDspPjtUCost 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6840
         TabIndex        =   27
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblDspPjtUCostL 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   9585
         TabIndex        =   26
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblPjtUCostL 
         Caption         =   "PJTUCOSTL"
         Height          =   255
         Left            =   8535
         TabIndex        =   25
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label lblPjtUCost 
         Caption         =   "PJTUCOST"
         Height          =   255
         Left            =   5775
         TabIndex        =   24
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label lblPjtUPriceL 
         Caption         =   "PJTUPRICEL"
         Height          =   255
         Left            =   2895
         TabIndex        =   23
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label lblPjtUPrice 
         Caption         =   "PJTUPRICE"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label lblPjtCurr 
         Caption         =   "CURR"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1040
         Width           =   1200
      End
      Begin VB.Label lblDspPjtUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   8520
         TabIndex        =   15
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblDspPjtUpdUsr 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   14
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblPjtUpdDate 
         Caption         =   "PJTUPDDATE"
         Height          =   240
         Left            =   6960
         TabIndex        =   13
         Top             =   6885
         Width           =   1500
      End
      Begin VB.Label lblPjtUpdUsr 
         Caption         =   "PJTUPDUSR"
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   6885
         Width           =   1500
      End
      Begin VB.Label lblDocNo 
         Caption         =   "PJTCODE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblPjtName 
         Caption         =   "PJTNAME"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   10080
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":9E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":A70A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":B436
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":B888
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":BBA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":BFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":C446
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":C760
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":CA7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":CECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":D7A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPjt001.frx":DAD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11460
      _ExtentX        =   20214
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
            Key             =   "SaveAs"
            ImageIndex      =   13
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   50
         EndProperty
      EndProperty
      BorderStyle     =   1
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
Attribute VB_Name = "frmPjt001"
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

Private Const LINENO = 0
Private Const ITMCODE = 1
Private Const BOOKNAME = 2
Private Const BOOKUNITPRICE = 3
Private Const BOOKQTY = 4
Private Const UOM = 5
Private Const Amt = 6
Private Const Amtl = 7
Private Const BOOKID = 8
Private Const SDUMMY = 9

Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcSaveAs = "SaveAs"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"

Private wiOpenDoc As Integer
Private wiAction As Integer

Private wlKey As Long
Private wlLineNo As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "MstProject"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wsOldCurCd As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, BOOKID
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
    
    wlKey = 0
    wlLineNo = 1
    wsOldCurCd = ""
    
    txtPjtUPrice = Format(0, gsAmtFmt)
    lblDspPjtUPriceL = Format(0, gsAmtFmt)
    lblDspPjtUCost = Format(0, gsAmtFmt)
    lblDspPjtUCostL = Format(0, gsAmtFmt)
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    FocusMe cboDocNo
End Sub

Private Sub Ini_Scr_AfrKey()
    If LoadRecord() = False Then
        wiAction = AddRec
        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
      
        wsOldCurCd = cboPjtCurr.Text
        
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    Call SetFieldStatus("AfrKey")
    
    txtPjtName.SetFocus
End Sub

Private Sub cboDocNo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo

    If gsLangID = "1" Then
        wsSQL = "SELECT PJTCODE, PJTENGNAME "
        wsSQL = wsSQL & " FROM MstProject "
        wsSQL = wsSQL & " WHERE PJTCODE LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
        wsSQL = wsSQL & " AND PJTSTATUS <> '2' "
        wsSQL = wsSQL & " ORDER BY PJTCODE "
    Else
        wsSQL = "SELECT PJTCODE, PJTCHINAME "
        wsSQL = wsSQL & " FROM MstProject "
        wsSQL = wsSQL & " WHERE PJTCODE LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
        wsSQL = wsSQL & " AND PJTSTATUS <> '2' "
        wsSQL = wsSQL & " ORDER BY PJTCODE "
    End If
    
    Call Ini_Combo(2, wsSQL, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLPJT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboDocNo_GotFocus()
    FocusMe cboDocNo
End Sub

Private Sub cboDocNo_LostFocus()
    FocusMe cboDocNo, True
End Sub

Private Sub cboDocNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNo, 15, KeyAscii)
    
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
        cboDocNo.SetFocus
        Exit Function
    End If
    
    If Chk_PJTCode(cboDocNo, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
   
            Exit Function
        End If
    End If
    
    Chk_cboDocNo = True
End Function

Private Sub cboPjtCurr_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboPjtCurr
    
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboPjtCurr.SelLength > 0, "", Set_Quote(cboPjtCurr.Text)) & "%' "
    wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Format(gsSystemDate, "MM")) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(gsSystemDate, "YYYY")) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
    Call Ini_Combo(2, wsSQL, cboPjtCurr.Left, cboPjtCurr.Top + cboPjtCurr.Height, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPjtCurr_GotFocus()
    FocusMe cboPjtCurr
End Sub

Private Sub cboPjtCurr_KeyPress(KeyAscii As Integer)
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim iCounter As Integer
    
    Call chk_InpLen(cboPjtCurr, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPjtCurr = False Then
            Exit Sub
        End If
        
        If getExcRate(cboPjtCurr.Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
            gsMsg = "沒有此貨幣!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboPjtCurr.SetFocus
            Exit Sub
        End If
        
        cboPjtTypeCode.SetFocus
    End If
End Sub

Private Sub cboPjtCurr_LostFocus()
    FocusMe cboPjtCurr, True
End Sub

Private Sub cboPjtTypeCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPjtTypeCode

    If gsLangID = "1" Then
        wsSQL = "SELECT ITMTYPECODE, ITMTYPECHIDESC "
        wsSQL = wsSQL & " FROM MstItemType"
        wsSQL = wsSQL & " WHERE ITMTYPECODE LIKE '%" & IIf(cboPjtTypeCode.SelLength > 0, "", Set_Quote(cboPjtTypeCode.Text)) & "%' "
        wsSQL = wsSQL & " AND ITMTYPESTATUS <> '2' "
        wsSQL = wsSQL & " ORDER BY ITMTYPECODE "
    Else
        wsSQL = "SELECT ITMTYPECODE, ITMTYPEENGDESC "
        wsSQL = wsSQL & " FROM MstItemType"
        wsSQL = wsSQL & " WHERE ITMTYPECODE LIKE '%" & IIf(cboPjtTypeCode.SelLength > 0, "", Set_Quote(cboPjtTypeCode.Text)) & "%' "
        wsSQL = wsSQL & " AND ITMTYPESTATUS <> '2' "
        wsSQL = wsSQL & " ORDER BY ITMTYPECODE "
    End If

    Call Ini_Combo(2, wsSQL, cboPjtTypeCode.Left, cboPjtTypeCode.Top + cboPjtTypeCode.Height, tblCommon, wsFormID, "TBLPJTTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPjtTypeCode_GotFocus()
    FocusMe cboPjtTypeCode
End Sub

Private Sub cboPjtTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPjtTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPjtTypeCode = False Then
            Exit Sub
        End If
        
        cboPjtUOMCode.SetFocus
    End If
End Sub

Private Sub cboPjtTypeCode_LostFocus()
    FocusMe cboPjtTypeCode, True
End Sub

Private Sub cboPjtUOMCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPjtUOMCode

    wsSQL = "SELECT UOMCODE, UOMDESC "
    wsSQL = wsSQL & " FROM MstUOM"
    wsSQL = wsSQL & " WHERE UOMCODE LIKE '%" & IIf(cboPjtUOMCode.SelLength > 0, "", Set_Quote(cboPjtUOMCode.Text)) & "%' "
    wsSQL = wsSQL & " AND UOMSTATUS <> '2' "
    wsSQL = wsSQL & " ORDER BY UOMSTATUS "

    Call Ini_Combo(2, wsSQL, cboPjtUOMCode.Left, cboPjtUOMCode.Top + cboPjtUOMCode.Height, tblCommon, wsFormID, "TBLPJTUOM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPjtUOMCode_GotFocus()
    FocusMe cboPjtUOMCode
End Sub

Private Sub cboPjtUOMCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPjtUOMCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPjtUOMCode = False Then
            Exit Sub
        End If
        
        txtPjtUPrice.SetFocus
    End If
End Sub

Private Sub cboPjtUOMCode_LostFocus()
    FocusMe cboPjtUOMCode, True
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
       Case vbKeyF6
            Call cmdOpen
        
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
            
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
       
        
        'Case vbKeyF3
        '    If wiAction = DefaultPage Then Call cmdDel
        
         'Case vbKeyF9
        
         '   If tbrProcess.Buttons(tcFind).Enabled = True Then Call cmdFind
            
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

Private Sub Ini_Caption()
On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblPjtUpdDate.Caption = Get_Caption(waScrItm, "PJTUPDDATE")
    lblPjtUpdUsr.Caption = Get_Caption(waScrItm, "PJTUPDUSR")
    lblPjtCurr.Caption = Get_Caption(waScrItm, "PJTCURR")
    lblPjtType.Caption = Get_Caption(waScrItm, "PJTTYPE")
    lblPjtName.Caption = Get_Caption(waScrItm, "PJTNAME")
    lblPjtUOMCode.Caption = Get_Caption(waScrItm, "PJTUOMCODE")
    lblPjtUPrice.Caption = Get_Caption(waScrItm, "PJTUPRICE")
    lblPjtUPriceL.Caption = Get_Caption(waScrItm, "PJTUPRICEL")
    lblPjtUCost.Caption = Get_Caption(waScrItm, "PJTUCOST")
    lblPjtUCostL.Caption = Get_Caption(waScrItm, "PJTUCOSTL")
        
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "ITMCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(BOOKUNITPRICE).Caption = Get_Caption(waScrItm, "BOOKUNITPRICE")
        .Columns(BOOKQTY).Caption = Get_Caption(waScrItm, "BOOKQTY")
        .Columns(UOM).Caption = Get_Caption(waScrItm, "UOM")
        .Columns(Amt).Caption = Get_Caption(waScrItm, "AMT")
        .Columns(Amtl).Caption = Get_Caption(waScrItm, "AMTL")
    End With
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcSaveAs).ToolTipText = Get_Caption(waScrToolTip, tcSaveAs)
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    lblKeyDesc = Get_Caption(waScrToolTip, "KEYDESC")
    lblComboPrompt = Get_Caption(waScrToolTip, "COMBOPROMPT")
    lblInsertLine = Get_Caption(waScrToolTip, "INSERTLINE")
    lblDeleteLine = Get_Caption(waScrToolTip, "DELETELINE")
    
    wsActNam(1) = Get_Caption(waScrItm, "PJTADD")
    wsActNam(2) = Get_Caption(waScrItm, "PJTEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PJTDELETE")
    
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

Private Sub Form_Unload(Cancel As Integer)
    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    
    Call UnLockAll(wsConnTime, wsFormID)
    
    Set waResult = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmPjt001 = Nothing
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If

End Sub

Private Function Chk_KeyFld() As Boolean
    Chk_KeyFld = False
    
    If Chk_cboPjtCurr = False Then
        Exit Function
    End If
    
    If Chk_cboPjtTypeCode = False Then
        Exit Function
    End If
    
    If Chk_cboPjtUOMCode = False Then
        Exit Function
    End If
    
    tblDetail.Enabled = True
    
    Chk_KeyFld = True
End Function

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
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
       Exit Function
    End If
    
    '' Last Check when Add
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    wlRowCtr = waResult.UpperBound(1)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_PJT001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, txtPjtName)
    Call SetSPPara(adcmdSave, 6, cboPjtTypeCode)
    Call SetSPPara(adcmdSave, 7, cboPjtCurr)
    Call SetSPPara(adcmdSave, 8, Format(txtPjtUPrice, gsAmtFmt))
    Call SetSPPara(adcmdSave, 9, Format(lblDspPjtUPriceL, gsAmtFmt))
    Call SetSPPara(adcmdSave, 10, Format(lblDspPjtUCost, gsAmtFmt))
    Call SetSPPara(adcmdSave, 11, Format(lblDspPjtUCostL, gsAmtFmt))
    Call SetSPPara(adcmdSave, 12, cboPjtUOMCode)
    Call SetSPPara(adcmdSave, 13, wsFormID)
    Call SetSPPara(adcmdSave, 14, gsUserID)
    Call SetSPPara(adcmdSave, 15, wsGenDte)
      
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 16)
    wsDocNo = GetSPPara(adcmdSave, 17)
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_PJT001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, wiCtr + 1)
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, BOOKNAME))
                Call SetSPPara(adcmdSave, 6, Format(waResult(wiCtr, BOOKUNITPRICE), gsAmtFmt))
                Call SetSPPara(adcmdSave, 7, Format(waResult(wiCtr, BOOKQTY), gsQtyFmt))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, UOM))
                Call SetSPPara(adcmdSave, 9, Format(waResult(wiCtr, Amt), gsAmtFmt))
                Call SetSPPara(adcmdSave, 10, Format(waResult(wiCtr, Amtl), gsAmtFmt))
                adcmdSave.Execute
            End If
        Next
    End If

    cnCon.CommitTrans
    
    If wiAction = AddRec Then
    If Trim(wsDocNo) <> "" Then
      '  Call cmdPrint(wsDocNo)
        gsMsg = "文件號 : " & wsDocNo & " 已製成!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    Else
        
        gsMsg = "文件儲存失敗!"
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

Private Function cmdSaveAs() As Boolean
    wiAction = AddRec
    Call GetNewKey
    Call cmdSave
End Function

Private Function cmdDel() As Boolean
    Dim wsGenDte As String
    Dim adcmdDelete As New ADODB.Command
    Dim wsDocNo As String
    Dim i As Integer
    
    cmdDel = False
    
    MousePointer = vbHourglass
    
    On Error GoTo cmdDelete_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Then
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
        
    adcmdDelete.CommandText = "USP_PJT001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, txtPjtName)
    Call SetSPPara(adcmdDelete, 6, "")
    Call SetSPPara(adcmdDelete, 7, "")
    Call SetSPPara(adcmdDelete, 8, "")
    Call SetSPPara(adcmdDelete, 9, "")
    Call SetSPPara(adcmdDelete, 10, "")
    Call SetSPPara(adcmdDelete, 11, "")
    Call SetSPPara(adcmdDelete, 12, "")
    Call SetSPPara(adcmdDelete, 13, wsFormID)
    Call SetSPPara(adcmdDelete, 14, gsUserID)
    Call SetSPPara(adcmdDelete, 15, wsGenDte)
      
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 16)
    wsDocNo = GetSPPara(adcmdDelete, 17)
    
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

Private Function InputValidation() As Boolean
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    If Not Chk_cboPjtCurr Then Exit Function
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, ITMCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = ITMCODE
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
            For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                    If waResult(wlCtr, ITMCODE) = waResult(wlCtr1, ITMCODE) Then
                      gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
                      MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                      tblDetail.Col = ITMCODE
                      tblDetail.SetFocus
                      Exit Function
                    End If
                End If
            Next
        Next
    End With
    
    
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tblDetail.Col = ITMCODE
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
        tblDetail.FirstRow = tblDetail.Row
        tblDetail.Col = ITMCODE
        tblDetail.SetFocus
        Exit Function
    End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cmdNew()
    Dim newForm As New frmPjt001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()
    Dim newForm As New frmPjt001
    
    newForm.OpenDoc = True
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show
End Sub

Private Sub Ini_Form()
    Me.KeyPreview = True
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Top = (Screen.Height - Me.Height) / 2
    
    Me.WindowState = 2
    'Me.tblDetail.Height = Me.Height - Me.tbrProcess.Height - Me.fra1.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "PJT001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "PJ"

End Sub

Private Sub cmdCancel()
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboDocNo.SetFocus
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

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
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
        Case tcSaveAs
            Call cmdSaveAs
        Case tcCancel
           If tbrProcess.Buttons(tcSave).Enabled = True Then
           If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, gsTitle) = vbYes Then
                Call cmdCancel
           End If
           Else
                Call cmdCancel
           End If
     '   Case tcFind
      '      Call cmdFind
        Case tcExit
            Unload Me
    End Select
End Sub

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
     '   .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = LINENO To BOOKID
            .Columns(wiCtr).AllowSizing = False
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
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                Case BOOKNAME
                    .Columns(wiCtr).Width = 3950
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case BOOKUNITPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case BOOKQTY
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case UOM
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).Button = True
                Case Amt
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Amtl
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case BOOKID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
                    .Columns(wiCtr).Locked = True
                    
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
Dim sTemp As String
   
    With tblDetail
        sTemp = .Columns(ColIndex)
        .Update
    End With


    If ColIndex = ITMCODE Then
        Call LoadBookGroup(sTemp)
    End If
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
Dim wsBookDefaultPrice As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsBookCurr As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case ITMCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdITMCODE(.Columns(ColIndex).Text, wsBookID, wsBookCode, wsBookName, wsBookDefaultPrice, wsBookCurr) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(LINENO).Text = wlLineNo
                .Columns(BOOKID).Text = wsBookID
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(BOOKUNITPRICE).Text = Format(wsBookDefaultPrice, gsAmtFmt)
                .Columns(BOOKQTY).Text = ""
                
                If Trim(.Columns(BOOKUNITPRICE).Text) <> "" Then
                    .Columns(Amt).Text = Format(To_Value(.Columns(BOOKUNITPRICE).Text) * To_Value(.Columns(BOOKQTY).Text), gsAmtFmt)
                End If
                
                If Trim(.Columns(BOOKUNITPRICE).Text) <> "" Then
                    .Columns(Amtl).Text = Format(To_Value(.Columns(BOOKUNITPRICE).Text) * To_Value(.Columns(BOOKQTY).Text), gsAmtFmt)
                End If
                
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(ColIndex).Text) <> wsBookCode Then
                    .Columns(ColIndex).Text = wsBookCode
                End If
            Case BOOKQTY, BOOKUNITPRICE
            
                If ColIndex = BOOKQTY Then
                    If Chk_grdBookQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                        
                ElseIf ColIndex = BOOKUNITPRICE Then
                    If Chk_grdBookUnitPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                End If
                    
                If Trim(.Columns(BOOKUNITPRICE).Text) <> "" Then
                    .Columns(Amt).Text = Format(To_Value(.Columns(BOOKUNITPRICE).Text) * To_Value(.Columns(BOOKQTY).Text), gsAmtFmt)
                End If
                
                If Trim(.Columns(BOOKUNITPRICE).Text) <> "" Then
                    .Columns(Amtl).Text = Format(To_Value(.Columns(BOOKUNITPRICE).Text) * To_Value(.Columns(BOOKQTY).Text), gsAmtFmt)
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
    
    On Error GoTo tblDetail_ButtonClick_Err
    With tblDetail
        Select Case ColIndex
            Case ITMCODE
                
                If gsLangID = 1 Then
                    wsSQL = "SELECT ITMCODE, ITMENGNAME ITNAME, STR(ITMDEFAULTPRICE,13,2) , FROM mstITEM "
                    wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSQL = wsSQL & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSQL = wsSQL & " '" & waResult(wiCtr, ITMCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                Else
                    wsSQL = "SELECT ITMCODE, ITMCHINAME ITNAME, STR(ITMDEFAULTPRICE,13,2) FROM mstITEM "
                    wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSQL = wsSQL & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSQL = wsSQL & " '" & waResult(wiCtr, ITMCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(3, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLITM", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
    
            Case UOM
                wsSQL = "SELECT UOMCODE, UOMDESC FROM MstUOM "
                wsSQL = wsSQL & " WHERE UOMSTATUS <> '2' AND UOMCODE LIKE '%" & Set_Quote(.Columns(UOM).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY UOMCODE "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLUOM", Me.Width, Me.Height)
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
                Case ITMCODE
                    KeyCode = vbDefault
                    .Col = BOOKNAME
                Case BOOKNAME
                    KeyCode = vbDefault
                    .Col = BOOKUNITPRICE
                Case BOOKUNITPRICE
                    KeyCode = vbDefault
                    .Col = BOOKQTY
                Case BOOKQTY
                    KeyCode = vbDefault
                    .Col = UOM
                Case UOM
                    KeyCode = vbDefault
                    .Col = Amt
                Case Amt
                    KeyCode = vbDefault
                    .Col = Amtl
                Case Amtl
                    KeyCode = vbKeyDown
                    .Col = ITMCODE
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
            Select Case .Col
                Case ITMCODE
                    .Col = ITMCODE
                Case Else
                    .Col = .Col - 1
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case Amtl
                    .Col = Amtl
                Case Else
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
        Case BOOKUNITPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
        
        Case BOOKQTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
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
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case ITMCODE
                    Call Chk_grdITMCODE(.Columns(ITMCODE).Text, "", "", "", "", "")
                'Case BOOKQTY
                '    Call Chk_grdBookQty(.Columns(BOOKQTY).Text)
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdITMCODE(inAccNo As String, outAccID As String, outAccNo As String, OutName As String, OutDefaultPrice As String, OutCurr As String) As Boolean
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
    
    wsSQL = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMDEFAULTPRICE, ITMCURR FROM MSTITEM"
    wsSQL = wsSQL & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        OutDefaultPrice = ReadRs(rsDes, "ITMDEFAULTPRICE")
        OutCurr = ReadRs(rsDes, "ITMCURR")
        
        Chk_grdITMCODE = True
    Else
        outAccID = ""
        OutName = ""
        OutDefaultPrice = ""
        OutCurr = ""
        gsMsg = "沒有物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
End Function

Private Function Chk_grdBookQty(inQty As String) As Boolean
    Chk_grdBookQty = False
    
    If Trim(inQty) = "" Then
        gsMsg = "沒有輸入物料數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If inQty < 1 Then
        gsMsg = "物料數量不可小於一本!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdBookQty = True
End Function

Private Function Chk_grdBookUnitPrice(inUnitPrice As String) As Boolean
    Chk_grdBookUnitPrice = False
    
    If inUnitPrice < 0 Then
        gsMsg = "價格不可小於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdBookUnitPrice = True
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
                   Trim(waResult(inRow, BOOKNAME)) = "" And _
                   Trim(waResult(inRow, BOOKUNITPRICE)) = "" And _
                   Trim(waResult(inRow, BOOKQTY)) = "" And _
                   Trim(waResult(inRow, BOOKID)) = "" Then
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
        
        If Chk_grdITMCODE(waResult(LastRow, ITMCODE), "", "", "", "", "") = False Then
            .Col = ITMCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdBookUnitPrice(waResult(LastRow, BOOKUNITPRICE)) = False Then
            .Col = BOOKUNITPRICE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdBookQty(waResult(LastRow, BOOKQTY)) = False Then
            .Col = BOOKQTY
            .Row = LastRow
            Exit Function
        End If
    
        If Chk_grdItmName(waResult(LastRow, BOOKNAME)) = False Then
            .Col = BOOKNAME
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdItmUOM(waResult(LastRow, UOM)) = False Then
            .Col = UOM
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdAmt(waResult(LastRow, Amt)) = False Then
            .Col = Amt
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdAmtL(waResult(LastRow, Amtl)) = False Then
            .Col = Amtl
            .Row = LastRow
            Exit Function
        End If
    
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
    SaveData = False
    
     If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) And _
        tbrProcess.Buttons(tcSave).Enabled = True Then
        
        gsMsg = "你是否確定不儲存現時之變更而離開?"
        If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbOK Then
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

Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = True
                .Buttons(tcEdit).Enabled = True
                .Buttons(tcDelete).Enabled = True
                .Buttons(tcSave).Enabled = False
                .Buttons(tcSaveAs).Enabled = False
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
                .Buttons(tcSaveAs).Enabled = False
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
                .Buttons(tcSaveAs).Enabled = False
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
                .Buttons(tcSaveAs).Enabled = True
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
                .Buttons(tcSaveAs).Enabled = True
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
                .Buttons(tcSaveAs).Enabled = True
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
            Me.txtPjtName.Enabled = False
            Me.cboPjtCurr.Enabled = False
            Me.cboPjtTypeCode.Enabled = False
            Me.cboPjtUOMCode.Enabled = False
            Me.txtPjtUPrice.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
        Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            Me.txtPjtName.Enabled = True
            Me.cboPjtCurr.Enabled = True
            Me.cboPjtTypeCode.Enabled = True
            Me.cboPjtUOMCode.Enabled = True
            Me.txtPjtUPrice.Enabled = True
            
            If wiAction <> AddRec Then
                Me.tblDetail.Enabled = True
            End If
    End Select
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoDup = False
    
        wsCurRec = tblDetail.Columns(ITMCODE)
 '       wsCurRecLn = tblDetail.Columns(wsWhsCode)
 
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, ITMCODE) Then
                  gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    If gsLangID = "1" Then
        wsSQL = "SELECT PJTID, PJTCODE, PJTTYPECODE, PJTENGNAME PJTNAME, "
        wsSQL = wsSQL & "PJTCURR, PJTUPRICE, PJTUPRICEL, PJTUCOST, PJTUCOSTL, "
        wsSQL = wsSQL & "PJTUOMCODE, PJTLASTUPD, PJTLASTUPDDATE, BOMID, BOMPJTID, "
        wsSQL = wsSQL & "BOMITMID, BOMNAME, BOMUPRICE, BOMQTY, BOMUOMCODE, "
        wsSQL = wsSQL & "BOMAMT, BOMAMTL, PJTLASTUPD, PJTLASTUPDDATE, BOMDTLN, "
        wsSQL = wsSQL & "ITMENGNAME ITMNAME, ITMCODE "
        wsSQL = wsSQL & "FROM  MSTPROJECT, MSTBOM, MSTITEM"
        wsSQL = wsSQL & "WHERE PJTCODE = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND PJTID = BOMPJTID "
        wsSQL = wsSQL & "AND ITMID = BOMITMID "
        wsSQL = wsSQL & "ORDER BY BOMDTLN "
    Else
        wsSQL = "SELECT PJTID, PJTCODE, PJTTYPECODE, PJTCHINAME PJTNAME, "
        wsSQL = wsSQL & "PJTCURR, PJTUPRICE, PJTUPRICEL, PJTUCOST, PJTUCOSTL, "
        wsSQL = wsSQL & "PJTUOMCODE, PJTLASTUPD, PJTLASTUPDDATE, BOMID, BOMPJTID, "
        wsSQL = wsSQL & "BOMITMID, BOMNAME, BOMUPRICE, BOMQTY, BOMUOMCODE, "
        wsSQL = wsSQL & "BOMAMT, BOMAMTL, PJTLASTUPD, PJTLASTUPDDATE, BOMDTLN, "
        wsSQL = wsSQL & "ITMCHINAME ITMNAME, ITMCODE "
        wsSQL = wsSQL & "FROM  MSTPROJECT, MSTBOM, MSTITEM "
        wsSQL = wsSQL & "WHERE PJTCODE = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND PJTID = BOMPJTID "
        wsSQL = wsSQL & "AND ITMID = BOMITMID "
        wsSQL = wsSQL & "ORDER BY BOMDTLN "
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsRcd, "PJTID")
    cboDocNo = ReadRs(rsRcd, "PJTCODE")
    txtPjtName = ReadRs(rsRcd, "PJTNAME")
    cboPjtCurr = ReadRs(rsRcd, "PJTCURR")
    cboPjtTypeCode = ReadRs(rsRcd, "PJTTYPECODE")
    cboPjtUOMCode = ReadRs(rsRcd, "PJTUOMCODE")
    txtPjtUPrice = ReadRs(rsRcd, "PJTUPRICE")
    lblDspPjtUPriceL = ReadRs(rsRcd, "PJTUPRICEL")
    lblDspPjtUCost = ReadRs(rsRcd, "PJTUCOST")
    lblDspPjtUCostL = ReadRs(rsRcd, "PJTUCOSTL")
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, BOOKID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsRcd, "BOMDTLN")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITMNAME")
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format(ReadRs(rsRcd, "BOMUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKQTY) = Format(ReadRs(rsRcd, "BOMQTY"), gsQtyFmt)
             waResult(.UpperBound(1), UOM) = ReadRs(rsRcd, "BOMUOMCODE")
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsRcd, "BOMAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsRcd, "BOMAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "BOMITMID")
             rsRcd.MoveNext
         Loop
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Call Calc_Total
    
    LoadRecord = True
End Function

Private Function Chk_KeyExist() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT PJTSTATUS FROM MSTPROJECT WHERE PJTCODE = '" & Set_Quote(cboDocNo) & "'"
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
        .TableKey = "PJTCODE"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function LoadBookGroup(ByVal ISBN As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wsITMID As String
    Dim wsISBN As String
    Dim wsName As String
    Dim wsBookDefaultPrice As String
    Dim wsBarCode As String
    Dim wsSeries As String
    Dim wsBookCurr As String
    
    Dim wsMtd As String
    
    LoadBookGroup = False
    wsMtd = ""
    
    If Trim(ISBN) = "" Then
        Exit Function
    End If
    
        wsSQL = "SELECT ITMSERIESNO "
        wsSQL = wsSQL & "FROM  mstITEM "
        wsSQL = wsSQL & "WHERE ITMSERIESNO = '" & Set_Quote(ISBN) & "' "
        wsSQL = wsSQL & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

        If rsRcd.RecordCount > 0 Then
            wsMtd = "1"
            wsSeries = ISBN
        End If
        rsRcd.Close
        Set rsRcd = Nothing
        
        If wsMtd <> "1" Then
    
        wsSQL = "SELECT ITMSERIESNO "
        wsSQL = wsSQL & "FROM  mstITEM "
        wsSQL = wsSQL & "WHERE ItmCode = '" & Set_Quote(ISBN) & "' "
        wsSQL = wsSQL & "AND ITMSERIESNO <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

        If rsRcd.RecordCount <= 0 Then
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
    
        If IsNull(ReadRs(rsRcd, "ITMSERIESNO")) Or Trim(ReadRs(rsRcd, "ITMSERIESNO")) = "" Then
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        Else
            wsSeries = ReadRs(rsRcd, "ITMSERIESNO")
        End If
    
         rsRcd.Close
         Set rsRcd = Nothing
       
         End If
         
    If gsLangID = "1" Then
        
        wsSQL = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
        wsSQL = wsSQL & "FROM  mstITEM "
        wsSQL = wsSQL & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSQL = wsSQL & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    Else

        wsSQL = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
        wsSQL = wsSQL & "FROM  mstITEM "
        wsSQL = wsSQL & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSQL = wsSQL & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
        
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If wsMtd = "" Then
    gsMsg = "此書為套裝書之一, 你是否要於此單表選擇全套書?"
    Else
    gsMsg = "此書為套裝書, 你是否要於此單表選擇全套書?"
    End If
    
    If MsgBox(gsMsg, vbInformation + vbYesNo, gsTitle) = vbNo Then
        tblDetail.ReBind
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If wsMtd = "1" Then
         With tblDetail
            .Delete
            .Update
            If .Row = -1 Then
                .Row = 0
            End If
         End With
    End If
    
    rsRcd.MoveFirst
    Do While Not rsRcd.EOF
  
       wsITMID = ReadRs(rsRcd, "ITMID")
       wsISBN = ReadRs(rsRcd, "ITMCODE")
       wsName = ReadRs(rsRcd, "ITNAME")
       wsBookDefaultPrice = ReadRs(rsRcd, "ITMDEFAULTPRICE")
       wsBarCode = ReadRs(rsRcd, "ITMBARCODE")
       wsBookCurr = ReadRs(rsRcd, "ITMCURR")
       
       With waResult
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), ITMCODE) = wsISBN
             waResult(.UpperBound(1), BOOKNAME) = wsName
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format(wsBookDefaultPrice, gsAmtFmt)
             waResult(.UpperBound(1), BOOKQTY) = 1
             waResult(.UpperBound(1), BOOKID) = wsITMID
             wlLineNo = wlLineNo + 1
        End With
    
     rsRcd.MoveNext
     Loop

    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    'Call Calc_Total
    
    LoadBookGroup = True
End Function

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

Public Function Chk_PJTCode(ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT PJTSTATUS FROM MSTPROJECT WHERE PJTCODE= '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "PJTSTATUS")
        Chk_PJTCode = True
    Else
        OutStatus = ""
        Chk_PJTCode = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub txtpjtname_GotFocus()
    FocusMe txtPjtName
End Sub

Private Sub txtpjtname_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPjtName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboPjtCurr.SetFocus
    End If
End Sub

Private Sub txtpjtname_LostFocus()
    FocusMe txtPjtName, True
End Sub

Private Function Chk_cboPjtCurr() As Boolean
    Chk_cboPjtCurr = False
     
    If Trim(cboPjtCurr.Text) = "" Then
        gsMsg = "必需輸入貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtCurr.SetFocus
        Exit Function
    End If
    
    
    If Chk_Curr(cboPjtCurr, gsSystemDate) = False Then
        gsMsg = "沒有此貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtCurr.SetFocus
       Exit Function
    End If
    
    Chk_cboPjtCurr = True
End Function

'Return calculated exchange price
Private Function getPrice(inFrmCurr As String, inToCurr As String, InExcr As String, inPrice As String) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wsPrice As String
    
    getPrice = ""
    
    wsSQL = "SELECT ExcRate FROM MstExchangeRate WHERE EXCCURR = '" & inFrmCurr & "' "
    wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Format(gsSystemDate, "MM")) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(gsSystemDate, "YYYY")) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsPrice = inPrice * ReadRs(rsRcd, "ExcRate")
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    getPrice = wsPrice / InExcr
End Function

Private Sub txtPjtUPrice_GotFocus()
    FocusMe txtPjtUPrice
End Sub

Private Sub txtPjtUPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPjtUPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPjtUPrice() = False Then
            Exit Sub
        Else
            Call Calc_Total
        End If
        
        If Chk_KeyFld() = True Then
            If tblDetail.Enabled = True Then
                tblDetail.Col = ITMCODE
                tblDetail.SetFocus
            Else
                txtPjtName.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtPjtUPrice_LostFocus()
    txtPjtUPrice = Format(txtPjtUPrice, gsAmtFmt)
    FocusMe txtPjtUPrice, True
End Sub

Private Function Chk_cboPjtTypeCode() As Boolean
    Chk_cboPjtTypeCode = False
    
    Dim wsDesc As String
     
    If Trim(cboPjtTypeCode) = "" Then
        gsMsg = "必需輸入工程類別!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtTypeCode.SetFocus
        Exit Function
    End If
    
    If Chk_PjtTypeCode(cboPjtTypeCode, wsDesc) = False Then
        gsMsg = "沒有此工程類別!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtTypeCode.SetFocus
       Exit Function
    End If
    
    Chk_cboPjtTypeCode = True
End Function

Public Function Chk_PjtTypeCode(ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT ITMTYPESTATUS FROM MSTITEMTYPE WHERE ITMTYPECODE= '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "ITMTYPESTATUS")
        Chk_PjtTypeCode = True
    Else
        OutStatus = ""
        Chk_PjtTypeCode = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboPjtUOMCode() As Boolean
    Chk_cboPjtUOMCode = False
    
    Dim wsDesc As String
     
    If Trim(cboPjtUOMCode) = "" Then
        gsMsg = "必需輸入單位!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtUOMCode.SetFocus
        Exit Function
    End If
    
    If Chk_PjtUOMCode(cboPjtUOMCode, wsDesc) = False Then
        gsMsg = "沒有此單位!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPjtUOMCode.SetFocus
       Exit Function
    End If
    
    Chk_cboPjtUOMCode = True
End Function

Public Function Chk_PjtUOMCode(ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT UOMSTATUS FROM MSTUOM WHERE UOMCODE= '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "UOMSTATUS")
        Chk_PjtUOMCode = True
    Else
        OutStatus = ""
        Chk_PjtUOMCode = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_grdItmName(inName As String) As Boolean
    Chk_grdItmName = False
    
    If Trim(inName) = "" Then
        gsMsg = "沒有輸入物料名稱!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdItmName = True
End Function

Private Function Chk_grdItmUOM(inUOM As String) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    
    Chk_grdItmUOM = False
    
    If Trim(inUOM) = "" Then
        gsMsg = "沒有輸入單位!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdItmUOM = False
        Exit Function
    End If
    
    wsSQL = "SELECT UOMCODE FROM MSTUOM"
    wsSQL = wsSQL & " WHERE UOMCODE = '" & Set_Quote(inUOM) & "'"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        Chk_grdItmUOM = True
    Else
        gsMsg = "輸入單位不正確!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdItmUOM = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
    Chk_grdItmUOM = True
End Function

Private Function Chk_grdAmt(inAmt As String) As Boolean
    Chk_grdAmt = False
    
    If Trim(inAmt) = "" Then
        gsMsg = "沒有輸入總額!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If inAmt < 1 Then
        gsMsg = "總額不可小於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdAmt = True
End Function

Private Function Chk_grdAmtL(inAmtL As String) As Boolean
    Chk_grdAmtL = False
    
    If Trim(inAmtL) = "" Then
        gsMsg = "沒有輸入總額!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If inAmtL < 1 Then
        gsMsg = "總額不可小於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdAmtL = True
End Function

Private Function Chk_txtPjtUPrice() As Boolean
    Chk_txtPjtUPrice = False
     
    If Trim(txtPjtUPrice.Text) = "" Then
        gsMsg = "必需輸入工程價錢!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPjtUPrice.SetFocus
        Exit Function
    End If
    
    If txtPjtUPrice.Text <= 0 Then
        gsMsg = "工程價錢必須大於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPjtUPrice.SetFocus
        Exit Function
    End If
    
    Chk_txtPjtUPrice = True
End Function

Private Function Calc_Total(Optional ByVal LastRow As Variant) As Boolean
    Dim wiTotalCost As Double
    Dim wiTotalCostL As Double
    Dim wsExcRate As String
    Dim wiRowCtr As Integer
    
    getExcRate cboPjtCurr.Text, Format(Date, "YYYY/MM/DD"), wsExcRate, ""
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalCost = wiTotalCost + To_Value(waResult(wiRowCtr, Amt))
    Next
    
    lblDspPjtUCost.Caption = Format(CStr(wiTotalCost), gsAmtFmt)
    lblDspPjtUCostL.Caption = Format(CStr(To_Value(wiTotalCost) * wsExcRate), gsAmtFmt)
    lblDspPjtUPriceL.Caption = Format(CStr(To_Value(txtPjtUPrice) * wsExcRate), gsAmtFmt)
    
    Calc_Total = True

End Function

