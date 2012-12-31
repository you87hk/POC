VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPC001 
   Caption         =   "Price Change Maintenance"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   11580
   Icon            =   "frmPC001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11040
      OleObjectBlob   =   "frmPC001.frx":030A
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5775
      Left            =   240
      OleObjectBlob   =   "frmPC001.frx":2A0D
      TabIndex        =   8
      Top             =   2160
      Width           =   11055
   End
   Begin VB.Frame fra1 
      Height          =   8055
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtPriChgDisPerIn 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame FRAPRICHGDOCTYPE 
         Height          =   615
         Left            =   6720
         TabIndex        =   18
         Top             =   960
         Width           =   4455
         Begin VB.OptionButton optPriChgDocType 
            Caption         =   "配售書目"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optPriChgDocType 
            Caption         =   "特別折扣"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPriChgRmk 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   9255
      End
      Begin MSMask.MaskEdBox medPriChgDocDate 
         Height          =   285
         Left            =   9840
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPriChgToDate 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPriChgFromDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblPriChgFromDate 
         Caption         =   "PRICHGFROMDATE"
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   1365
         Width           =   1890
      End
      Begin VB.Label lblPriChgToDate 
         Caption         =   "PRICHGTODATE"
         Height          =   225
         Left            =   3480
         TabIndex        =   22
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label lblPercent 
         Caption         =   "PERCENT"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lblPriChgDocDate 
         Caption         =   "PRICHGDOCDATE"
         Height          =   240
         Left            =   8640
         TabIndex        =   20
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label lblPriChgDisPerIn 
         Caption         =   "PRICHGDISPERIN"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   1800
      End
      Begin VB.Label lblDspPriChgUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   8520
         TabIndex        =   17
         Top             =   7560
         Width           =   2265
      End
      Begin VB.Label lblDspPriChgUpdUsr 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   16
         Top             =   7560
         Width           =   2265
      End
      Begin VB.Label lblPriChgUpdDate 
         Caption         =   "PRICHGUPDDATE"
         Height          =   240
         Left            =   6960
         TabIndex        =   15
         Top             =   7605
         Width           =   1500
      End
      Begin VB.Label lblPriChgUpdUsr 
         Caption         =   "PRICHGUPDUSR"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   7605
         Width           =   1500
      End
      Begin VB.Label lblDocNo 
         Caption         =   "DOCNO"
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
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblPriChgRmk 
         Caption         =   "REMARK"
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
         TabIndex        =   12
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":9C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":A4DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":ADB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":B20A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":B65C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":B976
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":BDC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":C21A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":C534
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":C84E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":CCA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPC001.frx":D57C
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
      Width           =   11580
      _ExtentX        =   20426
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmPC001"
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
Private wbReadOnly As Boolean

Private Const LINENO = 0
Private Const BOOKCODE = 1
Private Const BOOKNAME = 2
Private Const BOOKDEFAULTPRICE = 3
Private Const BOOKUNITPRICE = 4
Private Const BOOKCURR = 5
Private Const BOOKUOM = 6
Private Const BOOKID = 7
Private Const SDUMMY = 8

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

Private wlKey As Long
Private wlLineNo As Long
Private wsActNam(4) As String

Private wsPriChgDisPer As Long
Private wsConnTime As String
Private Const wsKeyType = "MstPriceChange"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String
Private wlRate As Long

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
    
    Call SetDateMask(medPriChgDocDate)
    Call SetDateMask(medPriChgFromDate)
    Call SetDateMask(medPriChgToDate)
    
    medPriChgDocDate.Text = gsSystemDate
    
    medPriChgFromDate = Left(gsSystemDate, 8) & "01"
    medPriChgToDate = Format(DateAdd("D", -1, DateAdd("M", 1, CDate(medPriChgFromDate.Text))), "YYYY/MM/DD")
    
    wbReadOnly = False
    wlLineNo = 1
    wlKey = 0
    txtPriChgDisPerIn = Format(0, gsAmtFmt)
    wsOldCurCd = ""
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    FocusMe cboDocNo
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim iCounter As Integer

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
      
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    For iCounter = 0 To optPriChgDocType.UBound
        If optPriChgDocType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Call SetFieldStatus("AfrKey")
    
    optPriChgDocType(iCounter).Value = True
    medPriChgDocDate.SetFocus
End Sub

Private Sub cboDocNo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSql = "SELECT PRICHGDOCNO, PRICHGRMK "
    wsSql = wsSql & " FROM MstPriceChange "
    wsSql = wsSql & " WHERE PRICHGDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND PRICHGSTATUS <> '2' "
    wsSql = wsSql & " ORDER BY PRICHGDOCNO "
    
    Call Ini_Combo(2, wsSql, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
    
    If Chk_PriChgDocNo(cboDocNo, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
   
            Exit Function
        End If
    End If
    
    Chk_cboDocNo = True
End Function

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
    lblPriChgRmk.Caption = Get_Caption(waScrItm, "PRICHGREMARK")
    lblPriChgUpdDate.Caption = Get_Caption(waScrItm, "PRICHGUPDDATE")
    lblPriChgUpdUsr.Caption = Get_Caption(waScrItm, "PRICHGUPDUSR")
    optPriChgDocType(0).Caption = Get_Caption(waScrItm, "PRICHGDOCTYPE1")
    optPriChgDocType(1).Caption = Get_Caption(waScrItm, "PRICHGDOCTYPE2")
    lblPriChgDocDate.Caption = Get_Caption(waScrItm, "PRICHGDOCDATE")
    lblPriChgDisPerIn.Caption = Get_Caption(waScrItm, "PRICHGDISPERIN")
    lblPercent.Caption = Get_Caption(waScrItm, "PERCENT")
    lblPriChgFromDate.Caption = Get_Caption(waScrItm, "FROMDATE")
    lblPriChgToDate.Caption = Get_Caption(waScrItm, "TODATE")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(BOOKDEFAULTPRICE).Caption = Get_Caption(waScrItm, "BOOKDEFAULTPRICE")
        .Columns(BOOKUNITPRICE).Caption = Get_Caption(waScrItm, "BOOKUNITPRICE")
        .Columns(BOOKCURR).Caption = Get_Caption(waScrItm, "BOOKCURR")
        .Columns(BOOKUOM).Caption = Get_Caption(waScrItm, "BOOKUOM")
    End With
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "PCADD")
    wsActNam(2) = Get_Caption(waScrItm, "PCEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PCDELETE")
    
    Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"
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
    Set frmPC001 = Nothing
End Sub

Private Sub medPriChgDocDate_GotFocus()
    FocusMe medPriChgDocDate
End Sub

Private Sub medPriChgDocDate_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, medPriChgDocDate, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_medPriChgDocDate = False Then
            Exit Sub
        End If
        
        txtPriChgRmk.SetFocus
    End If
End Sub

Private Sub medPriChgDocDate_LostFocus()
    FocusMe medPriChgDocDate, True
End Sub

Private Sub medPriChgFromDate_GotFocus()
    FocusMe medPriChgFromDate
End Sub

Private Sub medPriChgFromDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPriChgFromDate = False Then
            Call medPriChgFromDate_GotFocus
            Exit Sub
        End If
        
        If Trim(medPriChgFromDate) <> "/  /" And _
            Trim(medPriChgToDate) = "/  /" Then
            medPriChgToDate.Text = medPriChgFromDate.Text
        End If
        
        medPriChgToDate.SetFocus
    End If
End Sub

Private Function chk_medPriChgFromDate() As Boolean
    chk_medPriChgFromDate = False
    
    If Trim(medPriChgFromDate) = "/  /" Then
        gsMsg = "沒有輸入正確日期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgFromDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medPriChgFromDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgFromDate.SetFocus
        Exit Function
    End If
    
    chk_medPriChgFromDate = True
End Function

Private Function chk_medPriChgToDate() As Boolean
    chk_medPriChgToDate = False
    
    If UCase(medPriChgFromDate.Text) > UCase(medPriChgToDate.Text) Then
        gsMsg = "至到日期不可比由日期早!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgToDate.SetFocus
        Exit Function
    End If
    
    If Trim(medPriChgToDate) = "/  /" Then
        gsMsg = "沒有輸入正確日期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgToDate.SetFocus
        Exit Function
    End If

    If Chk_Date(medPriChgToDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgToDate.SetFocus
        Exit Function
    End If
    
    chk_medPriChgToDate = True
End Function

Private Sub medPriChgFromDate_LostFocus()
    FocusMe medPriChgFromDate, True
End Sub

Private Sub medPriChgToDate_GotFocus()
    FocusMe medPriChgToDate
End Sub

Private Sub medPriChgToDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPriChgToDate = False Then
            Call medPriChgToDate_GotFocus
            Exit Sub
        End If
        
        Call Opt_Setfocus(optPriChgDocType)
        
        wlRate = IIf(GetPriChgDocType = "A", 100 + txtPriChgDisPerIn, 100 - txtPriChgDisPerIn)
    End If
End Sub

Private Sub medPriChgToDate_LostFocus()
    FocusMe medPriChgToDate, True
End Sub

Private Sub optPriChgDocType_Click(Index As Integer)
    If Trim(txtPriChgDisPerIn) <> "" Then
        wlRate = IIf(GetPriChgDocType = "A", 100 + txtPriChgDisPerIn, 100 - txtPriChgDisPerIn)
    End If
End Sub

Private Sub optPriChgDocType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_optPriChgDocType() = True Then
            tblDetail.Col = BOOKCODE
            tblDetail.SetFocus
        End If
    End If
End Sub

Private Sub tblCommon_DblClick()
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case BOOKCODE
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
              Case BOOKCODE
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
    
    'If chk_txtItmLstExcr Then
    '    Exit Function
    'End If
    
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
        
    adcmdSave.CommandText = "USP_PC001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, GetPriChgDocType)
    Call SetSPPara(adcmdSave, 6, medPriChgDocDate)
    Call SetSPPara(adcmdSave, 7, medPriChgFromDate)
    Call SetSPPara(adcmdSave, 8, medPriChgToDate)
    Call SetSPPara(adcmdSave, 9, txtPriChgDisPerIn)
    Call SetSPPara(adcmdSave, 10, wlRate)
    Call SetSPPara(adcmdSave, 11, txtPriChgRmk)
    Call SetSPPara(adcmdSave, 12, wsFormID)
    Call SetSPPara(adcmdSave, 13, gsUserID)
    Call SetSPPara(adcmdSave, 14, wsGenDte)
      
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 15)
    wsDocNo = GetSPPara(adcmdSave, 16)
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_PC001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, wiCtr + 1)
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, BOOKCURR))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, BOOKUOM))
                Call SetSPPara(adcmdSave, 7, Format(waResult(wiCtr, BOOKDEFAULTPRICE), gsAmtFmt))
                Call SetSPPara(adcmdSave, 8, Format(waResult(wiCtr, BOOKUNITPRICE), gsAmtFmt))
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
        
    adcmdDelete.CommandText = "USP_PC001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, GetPriChgDocType)
    Call SetSPPara(adcmdDelete, 6, medPriChgDocDate)
    Call SetSPPara(adcmdDelete, 7, medPriChgFromDate)
    Call SetSPPara(adcmdDelete, 8, medPriChgToDate)
    Call SetSPPara(adcmdDelete, 9, txtPriChgDisPerIn)
    Call SetSPPara(adcmdDelete, 10, wlRate)
    Call SetSPPara(adcmdDelete, 11, txtPriChgRmk)
    Call SetSPPara(adcmdDelete, 12, wsFormID)
    Call SetSPPara(adcmdDelete, 13, gsUserID)
    Call SetSPPara(adcmdDelete, 14, wsGenDte)
      
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 15)
    wsDocNo = GetSPPara(adcmdDelete, 16)
    
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
    
    If Not Chk_txtPriChgDisPerIn Then Exit Function
    If Not Chk_medPriChgDocDate Then Exit Function
    If Not chk_medPriChgFromDate Then Exit Function
    If Not chk_medPriChgToDate Then Exit Function
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, BOOKCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = BOOKCODE
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
            For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                    If waResult(wlCtr, BOOKCODE) = waResult(wlCtr1, BOOKCODE) Then
                        gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
                        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                        tblDetail.Col = BOOKCODE
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
            tblDetail.Col = BOOKCODE
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
        tblDetail.FirstRow = tblDetail.Row
        tblDetail.Col = BOOKCODE
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
    Dim newForm As New frmPC001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()
    Dim newForm As New frmPC001
    
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
    wsFormID = "PC001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "PC"
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
        
        For wiCtr = LINENO To SDUMMY
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
                Case BOOKCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                Case BOOKNAME
                    .Columns(wiCtr).Width = 5000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case BOOKDEFAULTPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case BOOKUNITPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case BOOKCURR
                    .Columns(wiCtr).Width = 890
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Locked = True
                Case BOOKUOM
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = True
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


    If ColIndex = BOOKCODE Then
        Call LoadBookGroup(sTemp)
    End If
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsBookID As String
    Dim wsBookCode As String
    Dim wsBookName As String
    Dim wsBookDefaultPrice As String
    Dim wsPub As String
    Dim wdPrice As Double
    Dim wdDisPer As Double
    Dim wsBookCurr As String
    Dim wsBookUom As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case BOOKCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdBookCode(.Columns(ColIndex).Text, wsBookID, wsBookCode, "", wsBookName, wsBookDefaultPrice, wsBookCurr, wsBookUom) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(LINENO).Text = wlLineNo
                .Columns(BOOKID).Text = wsBookID
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(BOOKDEFAULTPRICE).Text = Format(wsBookDefaultPrice, gsAmtFmt)
                .Columns(BOOKUNITPRICE).Text = Format((wsBookDefaultPrice * wlRate) / 100, gsAmtFmt)
                .Columns(BOOKCURR).Text = wsBookCurr
                .Columns(BOOKUOM).Text = wsBookUom
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(ColIndex).Text) <> wsBookCode Then
                    .Columns(ColIndex).Text = wsBookCode
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
    Dim wsSql As String
    Dim wiTop As Long
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err
    With tblDetail
        Select Case ColIndex
            Case BOOKCODE
                
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSql = wsSql & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSql = wsSql & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(3, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLBOOKCODE", Me.Width, Me.Height)
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
            If .Bookmark = waResult.UpperBound(2) Then Exit Sub
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
                Case BOOKCODE, BOOKNAME, BOOKDEFAULTPRICE, BOOKUNITPRICE, BOOKCURR
                    KeyCode = vbDefault
                       .Col = .Col + 1
                Case BOOKUOM
                    KeyCode = vbKeyDown
                    .Col = BOOKCODE

            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case BOOKNAME, BOOKDEFAULTPRICE, BOOKUNITPRICE, BOOKCURR, BOOKUOM
                    .Col = .Col - 1
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case BOOKCODE, BOOKNAME, BOOKDEFAULTPRICE, BOOKUNITPRICE, BOOKCURR
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
        'Case Qty
        '    Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case BOOKUNITPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
    End Select
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = BOOKCODE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(BOOKCODE).Text, "", "", "", "", "", "", "")
                Case BOOKUNITPRICE
                    Call Chk_grdBookUnitPrice(.Columns(BOOKUNITPRICE).Text)
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdBookCode(inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, OutDefaultPrice As String, OutCurr As String, OutUOM As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入書號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
        Exit Function
    End If
    
    wsSql = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMBARCODE, ITMDEFAULTPRICE, ITMCURR, ITMUOMCODE FROM MSTITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        OutBarCode = ReadRs(rsDes, "ITMBARCODE")
        OutDefaultPrice = ReadRs(rsDes, "ITMDEFAULTPRICE")
        OutCurr = ReadRs(rsDes, "ITMCURR")
        OutUOM = ReadRs(rsDes, "ITMUOMCODE")
        
        Chk_grdBookCode = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        OutDefaultPrice = ""
        OutCurr = ""
        OutUOM = ""
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
End Function

Private Function Chk_grdBookQty(inQty As String) As Boolean
    Chk_grdBookQty = False
    
    If inQty < 1 Then
        gsMsg = "書本數量不可小於一本!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdBookQty = True
End Function

Private Function Chk_grdBookUnitPrice(inUnitPrice As String) As Boolean
    Chk_grdBookUnitPrice = False
    
    If Trim(inUnitPrice) = "" Then
        gsMsg = "沒有輸入價格!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
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
            If Trim(.Columns(BOOKCODE)) = "" Then
                Exit Function
            End If
        End With
    Else
        If waResult.UpperBound(1) >= 0 Then
            If Trim(waResult(inRow, BOOKCODE)) = "" And _
               Trim(waResult(inRow, BOOKNAME)) = "" And _
               Trim(waResult(inRow, BOOKDEFAULTPRICE)) = "" And _
               Trim(waResult(inRow, BOOKUNITPRICE)) = "" And _
               Trim(waResult(inRow, BOOKCURR)) = "" And _
               Trim(waResult(inRow, BOOKUOM)) = "" And _
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
        
        If Chk_grdBookCode(waResult(LastRow, BOOKCODE), "", "", "", "", "", "", "") = False Then
            .Col = BOOKCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdBookUnitPrice(waResult(LastRow, BOOKUNITPRICE)) = False Then
            .Col = BOOKUNITPRICE
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
Public Sub SetFieldStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
        Case "Default"
            Me.cboDocNo.Enabled = False
            
            Me.medPriChgDocDate.Enabled = False
            Me.medPriChgFromDate.Enabled = False
            Me.medPriChgToDate.Enabled = False
            Me.txtPriChgDisPerIn.Enabled = False
            Me.txtPriChgRmk.Enabled = False
            
            optPriChgDocType(0).Enabled = False
            optPriChgDocType(1).Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
        Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            optPriChgDocType(0).Enabled = True
            optPriChgDocType(1).Enabled = True
            
            Me.medPriChgDocDate.Enabled = True
            Me.medPriChgFromDate.Enabled = True
            Me.medPriChgToDate.Enabled = True
            Me.txtPriChgDisPerIn.Enabled = True
            Me.txtPriChgRmk.Enabled = True
            Me.tblDetail.Enabled = True
    End Select
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(BOOKCODE)
 '       wsCurRecLn = tblDetail.Columns(wsWhsCode)
 
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           If wsCurRec = waResult(wlCtr, BOOKCODE) Then
              gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
              MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
              Exit Function
           End If
        End If
    Next
    
    Chk_NoDup = True
End Function

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    wsSql = "SELECT PRICHGDOCID, PRICHGDOCNO, PRICHGRMK, PRICHGDOCTYPE, PRICHGDOCDATE, PRICHGDISPERIN, PRICHGDTITEMID, ITMCODE, PRICHGDTREMARK, ITMBARCODE, ITMCHINAME, PRICHGDTDEFAULTPRICE, PRICHGDTUNITPRICE, PRICHGDTCURR, PRICHGDTUOM, PRICHGDTDOCLINE, "
    wsSql = wsSql & "PRICHGFROMDATE, PRICHGTODATE "
    wsSql = wsSql & "FROM MSTPRICECHANGE, MSTPRICECHANGEDT, MSTITEM "
    wsSql = wsSql & "WHERE PRICHGDOCNO = '" & cboDocNo & "' "
    wsSql = wsSql & "AND PRICHGDTDOCID = PRICHGDOCID "
    wsSql = wsSql & "AND ITMID = PRICHGDTITEMID "
    wsSql = wsSql & "ORDER BY PRICHGDTDOCLINE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsRcd, "PRICHGDOCID")
    txtPriChgRmk = ReadRs(rsRcd, "PRICHGRMK")
    txtPriChgDisPerIn = Format(ReadRs(rsRcd, "PRICHGDISPERIN"), gsAmtFmt)
    medPriChgDocDate = Dsp_MedDate(ReadRs(rsRcd, "PRICHGDOCDATE"))
    medPriChgFromDate = Dsp_MedDate(ReadRs(rsRcd, "PRICHGFROMDATE"))
    medPriChgToDate = Dsp_MedDate(ReadRs(rsRcd, "PRICHGTODATE"))
    
    SetPriChgDocType ReadRs(rsRcd, "PRICHGDOCTYPE")
        
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, BOOKID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsRcd, "PRICHGDTDOCLINE")
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITMCHINAME")
             waResult(.UpperBound(1), BOOKDEFAULTPRICE) = Format(ReadRs(rsRcd, "PRICHGDTDEFAULTPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format(ReadRs(rsRcd, "PRICHGDTUNITPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKCURR) = ReadRs(rsRcd, "PRICHGDTCURR")
             waResult(.UpperBound(1), BOOKUOM) = ReadRs(rsRcd, "PRICHGDTUOM")
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "PRICHGDTITEMID")
             rsRcd.MoveNext
         Loop
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    LoadRecord = True
End Function

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT PRICHGSTATUS FROM MstPriceChange WHERE PRICHGDOCNO = '" & Set_Quote(cboDocNo) & "'"
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
        .TableKey = "PriChgDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function LoadBookGroup(ByVal ISBN As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wsItmID As String
    Dim wsISBN As String
    Dim wsName As String
    Dim wsBookDefaultPrice As String
    Dim wsBarCode As String
    Dim wsSeries As String
    Dim wsBookCurr As String
    Dim wsBookUom As String
    
    Dim wsMtd As String
    
    LoadBookGroup = False
    wsMtd = ""
    
    If Trim(ISBN) = "" Then
        Exit Function
    End If
    
        wsSql = "SELECT ITMSERIESNO "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ITMSERIESNO = '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

        If rsRcd.RecordCount > 0 Then
            wsMtd = "1"
            wsSeries = ISBN
        End If
        rsRcd.Close
        Set rsRcd = Nothing
        
        If wsMtd <> "1" Then
    
        wsSql = "SELECT ITMSERIESNO "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmCode = '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "AND ITMSERIESNO <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

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
        
        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE, ITMUOMCODE "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "ORDER BY ItmCode "
    Else

        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE, ITMUOMCODE "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "ORDER BY ItmCode "
        
    End If
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

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
  
       wsItmID = ReadRs(rsRcd, "ITMID")
       wsISBN = ReadRs(rsRcd, "ITMCODE")
       wsName = ReadRs(rsRcd, "ITNAME")
       wsBookDefaultPrice = ReadRs(rsRcd, "ITMDEFAULTPRICE")
       wsBarCode = ReadRs(rsRcd, "ITMBARCODE")
       wsBookCurr = ReadRs(rsRcd, "ITMCURR")
       wsBookUom = ReadRs(rsRcd, "ITMUOMCODE")
       
       With waResult
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), BOOKCODE) = wsISBN
             waResult(.UpperBound(1), BOOKNAME) = wsName
             waResult(.UpperBound(1), BOOKDEFAULTPRICE) = Format(wsBookDefaultPrice, gsAmtFmt)
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format((wsBookDefaultPrice * wlRate) / 100, gsAmtFmt)
             waResult(.UpperBound(1), BOOKCURR) = wsBookCurr
             waResult(.UpperBound(1), BOOKUOM) = wsBookUom
             waResult(.UpperBound(1), BOOKID) = wsItmID
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
            
            If .Bookmark = waResult.UpperBound(2) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
End Sub

Public Function Chk_PriChgDocNo(ByVal inDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT PRICHGSTATUS FROM MSTPRICECHANGE WHERE PRICHGDOCNO= '" & Set_Quote(inDocNo) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "PRICHGSTATUS")
        Chk_PriChgDocNo = True
    Else
        OutStatus = ""
        Chk_PriChgDocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub txtPriChgDisPerIn_GotFocus()
    FocusMe txtPriChgDisPerIn
End Sub

Private Sub txtPriChgDisPerIn_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPriChgDisPerIn, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPriChgDisPerIn = False Then
            Exit Sub
        End If
        
        medPriChgFromDate.SetFocus
    End If
End Sub

Private Sub txtPriChgDisPerIn_LostFocus()
    FocusMe txtPriChgDisPerIn, True
    txtPriChgDisPerIn = Format(txtPriChgDisPerIn.Text, gsAmtFmt)
End Sub

Private Sub txtPriChgRmk_GotFocus()
    FocusMe txtPriChgRmk
End Sub

Private Sub txtPriChgRmk_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPriChgRmk, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtPriChgDisPerIn.SetFocus
    End If
End Sub

Private Sub txtPriChgRmk_LostFocus()
    FocusMe txtPriChgRmk, True
End Sub

Private Sub SetPriChgDocType(ByVal inCode As String)
    Select Case inCode
        Case "A"
            optPriChgDocType(0).Value = True
            
        Case "D"
            optPriChgDocType(1).Value = True
            
    End Select
End Sub

Private Function GetPriChgDocType() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To optPriChgDocType.UBound
        If optPriChgDocType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetPriChgDocType = "A"
            
        Case 1
            GetPriChgDocType = "D"
    End Select
End Function

'Return calculated exchange price
Private Function getPrice(inFrmCurr As String, inToCurr As String, inExcr As String, inPrice As String) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wsPrice As String
    
    getPrice = ""
    
    wsSql = "SELECT ExcRate FROM MstExchangeRate WHERE EXCCURR = '" & inFrmCurr & "' "
    wsSql = wsSql & " AND EXCMN = '" & To_Value(Format(gsSystemDate, "MM")) & "' "
    wsSql = wsSql & " AND EXCYR = '" & Set_Quote(Format(gsSystemDate, "YYYY")) & "' "
    wsSql = wsSql & " AND EXCSTATUS = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsPrice = inPrice * ReadRs(rsRcd, "ExcRate")
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    getPrice = wsPrice / inExcr
End Function

Private Function Chk_medPriChgDocDate() As Boolean
    Chk_medPriChgDocDate = False
    
    If Trim(medPriChgDocDate.Text) = "/" Then
        Chk_medPriChgDocDate = True
        Exit Function
    End If
    
    If Chk_Date(medPriChgDocDate) = False Then
        gsMsg = "日子不正確!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPriChgDocDate.SetFocus
        Exit Function
    End If
    
    Chk_medPriChgDocDate = True
End Function

Private Function Chk_txtPriChgDisPerIn() As Boolean
    Chk_txtPriChgDisPerIn = False
    
    If Trim(txtPriChgDisPerIn) = "" Then
        gsMsg = "沒有輸入正確值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    If txtPriChgDisPerIn < 0 Then
        gsMsg = "數值不可少於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    If txtPriChgDisPerIn > 100 Then
        gsMsg = "數值不可多於一百!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    Chk_txtPriChgDisPerIn = True
End Function

Private Function chk_optPriChgDocType() As Boolean
    chk_optPriChgDocType = False
    
    If Trim(txtPriChgDisPerIn) = "" Then
        gsMsg = "沒有輸入正確值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    If txtPriChgDisPerIn < 0 Then
        gsMsg = "數值不可少於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    If txtPriChgDisPerIn > 100 Then
        gsMsg = "數值不可多於一百!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPriChgDisPerIn.SetFocus
        Exit Function
    End If
    
    wlRate = IIf(GetPriChgDocType = "A", 100 + txtPriChgDisPerIn, 100 - txtPriChgDisPerIn)
    
    chk_optPriChgDocType = True
End Function
