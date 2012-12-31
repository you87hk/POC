VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmBL001 
   Caption         =   "Book List Maintenance"
   ClientHeight    =   7920
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   11460
   Icon            =   "frmBL001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11460
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11040
      OleObjectBlob   =   "frmBL001.frx":030A
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   4935
      Left            =   240
      OleObjectBlob   =   "frmBL001.frx":2A0D
      TabIndex        =   23
      Top             =   2280
      Width           =   11055
   End
   Begin VB.ComboBox cboItmLstCurr 
      Height          =   300
      Left            =   2040
      TabIndex        =   17
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame fra1 
      Height          =   7335
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   11295
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   5040
         TabIndex        =   18
         Top             =   120
         Width           =   6135
         Begin VB.Label lblKeyDesc 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   360
            TabIndex        =   22
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblComboPrompt 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   1920
            TabIndex        =   21
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblInsertLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   3360
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblDeleteLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   4800
            TabIndex        =   19
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.TextBox txtItmLstExcr 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Frame FRADOCTYPE 
         Height          =   615
         Left            =   6720
         TabIndex        =   14
         Top             =   960
         Width           =   4455
         Begin VB.OptionButton optDocType 
            Caption         =   "配售書目"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optDocType 
            Caption         =   "特別折扣"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   9255
      End
      Begin VB.Label lblItmLstCurr 
         Caption         =   "CURR"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1040
         Width           =   1200
      End
      Begin VB.Label lblItmLstExcr 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label lblDspItmLstUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   8520
         TabIndex        =   13
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblDspItmLstUpdUsr 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblItmLstUpdDate 
         Caption         =   "ITMLSTUPDDATE"
         Height          =   240
         Left            =   6960
         TabIndex        =   11
         Top             =   6885
         Width           =   1500
      End
      Begin VB.Label lblItmLstUpdUsr 
         Caption         =   "ITMLSTUPDUSR"
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   6885
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
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblRemark 
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
         TabIndex        =   8
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
            Picture         =   "frmBL001.frx":8CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":95CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":9EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":A2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":A748
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":AA62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":AEB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":B306
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":B620
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":B93A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":BD8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":C668
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBL001.frx":C990
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
Attribute VB_Name = "frmBL001"
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
Private Const BOOKCODE = 1
Private Const BARCODE = 2
Private Const BOOKNAME = 3
Private Const BOOKUNITPRICE = 4
Private Const BOOKQTY = 5
Private Const BOOKID = 6
Private Const SDUMMY = 7

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
Private Const wsKeyType = "MstItemList"
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
      
        wsOldCurCd = cboItmLstCurr.Text
        
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    For iCounter = 0 To optDocType.UBound
        If optDocType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Call SetFieldStatus("AfrKey")
    
    If UCase(cboItmLstCurr) = UCase(wsBaseCurCd) Then
        txtItmLstExcr.Text = Format("1", gsExrFmt)
        txtItmLstExcr.Enabled = False
    Else
        txtItmLstExcr.Enabled = True
    End If
    
    optDocType(iCounter).Value = True
    txtRemark.SetFocus
End Sub

Private Sub cboDocNo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSql = "SELECT ITMLSTDOCNO, ITMLSTRMK "
    wsSql = wsSql & " FROM MstItemList "
    wsSql = wsSql & " WHERE ITMLSTDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND ITMLSTSTATUS <> '2' "
    wsSql = wsSql & " ORDER BY ITMLSTDOCNO "
    
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
    
    If Chk_ItmLstDocNo(cboDocNo, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
   
            Exit Function
        End If
    End If
    
    Chk_cboDocNo = True
End Function

Private Sub cboItmLstCurr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmLstCurr
    
    
    wsSql = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboItmLstCurr.SelLength > 0, "", Set_Quote(cboItmLstCurr.Text)) & "%' "
    wsSql = wsSql & " AND EXCMN = '" & To_Value(Format(gsSystemDate, "MM")) & "' "
    wsSql = wsSql & " AND EXCYR = '" & Set_Quote(Format(gsSystemDate, "YYYY")) & "' "
    wsSql = wsSql & " AND EXCSTATUS = '1' "
    wsSql = wsSql & "ORDER BY EXCCURR "
    Call Ini_Combo(2, wsSql, cboItmLstCurr.Left, cboItmLstCurr.Top + cboItmLstCurr.Height, tblCommon, "BL001", "TBLCURCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmLstCurr_GotFocus()
    FocusMe cboItmLstCurr
End Sub

Private Sub cboItmLstCurr_KeyPress(KeyAscii As Integer)
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim iCounter As Integer
    
    Call chk_InpLen(cboItmLstCurr, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmLstCurr = False Then
            Exit Sub
        End If
        
        If getExcRate(cboItmLstCurr.Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
            gsMsg = "沒有此貨幣!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            txtItmLstExcr.Text = Format(0, gsExrFmt)
            cboItmLstCurr.SetFocus
            Exit Sub
        End If
        
        If wsOldCurCd <> cboItmLstCurr.Text Then
            txtItmLstExcr.Text = Format(wsExcRate, gsExrFmt)
            wsOldCurCd = cboItmLstCurr.Text
        End If
        
        If UCase(cboItmLstCurr) = UCase(wsBaseCurCd) Then
            txtItmLstExcr.Text = Format("1", gsExrFmt)
            txtItmLstExcr.Enabled = False
        Else
            txtItmLstExcr.Enabled = True
        End If
        
        If txtItmLstExcr.Enabled Then
            txtItmLstExcr.SetFocus
        Else
            For iCounter = 0 To optDocType.UBound
                If optDocType(iCounter).Value = True Then
                    optDocType(iCounter).SetFocus
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub cboItmLstCurr_LostFocus()
    FocusMe cboItmLstCurr, True
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
    lblRemark.Caption = Get_Caption(waScrItm, "REMARK")
    lblItmLstUpdDate.Caption = Get_Caption(waScrItm, "ITMLSTUPDDATE")
    lblItmLstUpdUsr.Caption = Get_Caption(waScrItm, "ITMLSTUPDUSR")
    optDocType(0).Caption = Get_Caption(waScrItm, "ITMLSTDOCTYPE1")
    optDocType(1).Caption = Get_Caption(waScrItm, "ITMLSTDOCTYPE2")
    lblItmLstCurr.Caption = Get_Caption(waScrItm, "ITMLSTCURR")
    lblItmLstExcr.Caption = Get_Caption(waScrItm, "ITMLSTEXCR")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(BOOKUNITPRICE).Caption = Get_Caption(waScrItm, "BOOKUNITPRICE")
        .Columns(BOOKQTY).Caption = Get_Caption(waScrItm, "BOOKQTY")
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
    
    wsActNam(1) = Get_Caption(waScrItm, "BLADD")
    wsActNam(2) = Get_Caption(waScrItm, "BLEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "BLDELETE")
    
    Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
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
    Set frmBL001 = Nothing
End Sub

Private Sub optDocType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        tblDetail.SetFocus
        tblDetail.Col = BOOKCODE
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
    
    If Chk_cboItmLstCurr Then
        Exit Function
    End If
    
    If chk_txtItmLstExcr Then
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
        
    adcmdSave.CommandText = "USP_BL001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, txtRemark)
    Call SetSPPara(adcmdSave, 6, wsFormID)
    Call SetSPPara(adcmdSave, 7, GetItmLstDocType)
    Call SetSPPara(adcmdSave, 8, cboItmLstCurr)
    Call SetSPPara(adcmdSave, 9, txtItmLstExcr)
    Call SetSPPara(adcmdSave, 10, gsUserID)
    Call SetSPPara(adcmdSave, 11, wsGenDte)
      
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 12)
    wsDocNo = GetSPPara(adcmdSave, 13)
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_BL001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, wiCtr + 1)
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, BOOKQTY))
                Call SetSPPara(adcmdSave, 6, Format(waResult(wiCtr, BOOKUNITPRICE), gsAmtFmt))
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
        
    adcmdDelete.CommandText = "USP_BL001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, txtRemark)
    Call SetSPPara(adcmdDelete, 6, wsFormID)
    Call SetSPPara(adcmdDelete, 7, GetItmLstDocType)
    Call SetSPPara(adcmdDelete, 8, cboItmLstCurr)
    Call SetSPPara(adcmdDelete, 9, txtItmLstExcr)
    Call SetSPPara(adcmdDelete, 10, gsUserID)
    Call SetSPPara(adcmdDelete, 11, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 12)
    wsDocNo = GetSPPara(adcmdDelete, 13)
    
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
    
    If Not Chk_cboItmLstCurr Then Exit Function
    If Not chk_txtItmLstExcr Then Exit Function
    
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
    Dim newForm As New frmBL001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()
    Dim newForm As New frmBL001
    
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
    wsFormID = "BL001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "BL"

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
                Case BOOKCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                Case BARCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case BOOKNAME
                    .Columns(wiCtr).Width = 5500
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case BOOKUNITPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case BOOKQTY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Locked = False
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
            Case BOOKCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdBookCode(.Columns(ColIndex).Text, wsBookID, wsBookCode, wsBarCode, wsBookName, wsBookDefaultPrice, wsBookCurr) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(LINENO).Text = wlLineNo
                .Columns(BOOKID).Text = wsBookID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(BOOKUNITPRICE).Text = Format(getPrice(wsBookCurr, cboItmLstCurr, txtItmLstExcr, wsBookDefaultPrice), gsAmtFmt)
                .Columns(BOOKQTY).Text = ""
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
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, STR(ITMDEFAULTPRICE,13,2) , FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSql = wsSql & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, STR(ITMDEFAULTPRICE,13,2) FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSql = wsSql & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(4, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLBOOKCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
          '  Case WhsCode
                
          '      wsSql = "SELECT WHSCODE, WHSDESC FROM mstWareHouse "
          '      wsSql = wsSql & " WHERE WHSSTATUS <> '2' AND WHSCODE LIKE '%" & Set_Quote(.Columns(WhsCode).Text) & "%' "
          '      wsSql = wsSql & " ORDER BY WHSCODE "
                
          '      Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
          '      tblCommon.Visible = True
          '      tblCommon.SetFocus
          '      Set wcCombo = tblDetail
                
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
                Case BOOKCODE
                    KeyCode = vbDefault
                    .Col = BOOKUNITPRICE
                Case BOOKUNITPRICE
                    KeyCode = vbDefault
                    .Col = BOOKQTY
                Case BOOKQTY
                    KeyCode = vbKeyDown
                    .Col = BOOKCODE
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case BOOKUNITPRICE
                    .Col = BOOKCODE
                Case BOOKQTY
                    .Col = BOOKUNITPRICE
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case BOOKCODE
                    .Col = BOOKUNITPRICE
                Case BOOKUNITPRICE
                    .Col = BOOKQTY
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
           .Col = BOOKCODE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(BOOKCODE).Text, "", "", "", "", "", "")
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

Private Function Chk_grdBookCode(inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, OutDefaultPrice As String, OutCurr As String) As Boolean
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
    
    
    wsSql = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMBARCODE, ITMDEFAULTPRICE, ITMCURR FROM MSTITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        OutBarCode = ReadRs(rsDes, "ITMBARCODE")
        OutDefaultPrice = ReadRs(rsDes, "ITMDEFAULTPRICE")
        OutCurr = ReadRs(rsDes, "ITMCURR")
        
        Chk_grdBookCode = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        OutDefaultPrice = ""
        OutCurr = ""
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
End Function

Private Function Chk_grdBookQty(inQty As String) As Boolean
    Chk_grdBookQty = False
    
    If Trim(inQty) = "" Then
        gsMsg = "沒有輸入書本數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If inQty < 1 Then
        gsMsg = "書本數量不可小於一本!"
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
                If Trim(.Columns(BOOKCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, BOOKCODE)) = "" And _
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
        
        If Chk_grdBookCode(waResult(LastRow, BOOKCODE), "", "", "", "", "", "") = False Then
            .Col = BOOKCODE
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
Public Sub SetFieldStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
        Case "Default"
            Me.cboDocNo.Enabled = False
            Me.txtRemark.Enabled = False
            optDocType(0).Enabled = False
            optDocType(1).Enabled = False
            Me.cboItmLstCurr.Enabled = False
            Me.txtItmLstExcr.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
           Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
           Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            optDocType(0).Enabled = True
            optDocType(1).Enabled = True
            Me.cboItmLstCurr.Enabled = True
            Me.txtItmLstExcr.Enabled = True
            Me.txtRemark.Enabled = True
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
    
    wsSql = "SELECT ITMLSTDOCID, ITMLSTDOCNO, ITMLSTRMK, ITMLSTDOCTYPE, ITMLSTCURR, ITMLSTEXCR, ITMLSTDTITEMID, ITMLSTDTDOCLINE, ITMCODE, ITMLSTDTQTY, ITMBARCODE, ITMCHINAME, ITMLSTDTUNITPRICE "
    wsSql = wsSql & "FROM  MSTITEMLIST, MSTITEMLISTDT, MSTITEM "
    wsSql = wsSql & "WHERE ITMLSTDOCNO = '" & cboDocNo & "' "
    wsSql = wsSql & "AND ITMLSTDTDOCID = ITMLSTDOCID "
    wsSql = wsSql & "AND ITMID = ITMLSTDTITEMID "
    wsSql = wsSql & "ORDER BY ITMLSTDTDOCLINE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsRcd, "ITMLSTDOCID")
    txtRemark = ReadRs(rsRcd, "ITMLSTRMK")
    txtItmLstExcr = Format(ReadRs(rsRcd, "ITMLSTEXCR"), gsExrFmt)
    cboItmLstCurr = ReadRs(rsRcd, "ITMLSTCURR")
    
    SetItmLstDocType ReadRs(rsRcd, "ITMLSTDOCTYPE")
        
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, BOOKID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsRcd, "ITMLSTDTDOCLINE")
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITMCHINAME")
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format(ReadRs(rsRcd, "ITMLSTDTUNITPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKQTY) = Format(ReadRs(rsRcd, "ITMLSTDTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "ITMLSTDTITEMID")
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

    
    wsSql = "SELECT ITMLSTSTATUS FROM MSTITEMLIST WHERE ITMLSTDOCNO = '" & Set_Quote(cboDocNo) & "'"
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
        .TableKey = "ItmLstDocNo"
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
        
        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "ORDER BY ItmCode "
    Else

        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
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
       
       With waResult
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), BOOKCODE) = wsISBN
             waResult(.UpperBound(1), BARCODE) = wsBarCode
             waResult(.UpperBound(1), BOOKNAME) = wsName
             waResult(.UpperBound(1), BOOKUNITPRICE) = Format(getPrice(wsBookCurr, cboItmLstCurr, txtItmLstExcr, wsBookDefaultPrice), gsAmtFmt)
             waResult(.UpperBound(1), BOOKQTY) = 1
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

Public Function Chk_ItmLstDocNo(ByVal inDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT ITMLSTSTATUS FROM MSTITEMLIST WHERE ITMLSTDOCNO= '" & Set_Quote(inDocNo) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "ITMLSTSTATUS")
        Chk_ItmLstDocNo = True
    Else
        OutStatus = ""
        Chk_ItmLstDocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub txtItmLstExcr_GotFocus()
    FocusMe txtItmLstExcr
End Sub

Private Sub txtItmLstExcr_KeyPress(KeyAscii As Integer)
    Dim iCounter As Integer

    Call Chk_InpNum(KeyAscii, txtItmLstExcr.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtItmLstExcr Then
            If Chk_KeyFld Then
                For iCounter = 0 To optDocType.UBound
                    If optDocType(iCounter).Value = True Then
                        optDocType(iCounter).SetFocus
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub txtItmLstExcr_LostFocus()
    FocusMe txtItmLstExcr, True
End Sub

Private Sub txtRemark_GotFocus()
    FocusMe txtRemark
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtRemark, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboItmLstCurr.SetFocus
    End If
End Sub

Private Sub txtRemark_LostFocus()
    FocusMe txtRemark, True
End Sub

Private Sub SetItmLstDocType(ByVal inCode As String)
    Select Case inCode
        Case "B"
            optDocType(0).Value = True
            
        Case "D"
            optDocType(1).Value = True
            
    End Select
End Sub

Private Function GetItmLstDocType() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To optDocType.UBound
        If optDocType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetItmLstDocType = "B"
            
        Case 1
            GetItmLstDocType = "D"
    End Select
End Function

Private Function Chk_cboItmLstCurr() As Boolean
    Chk_cboItmLstCurr = False
     
    If Trim(cboItmLstCurr.Text) = "" Then
        gsMsg = "必需輸入貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmLstCurr.SetFocus
        Exit Function
    End If
    
    
    If Chk_Curr(cboItmLstCurr, gsSystemDate) = False Then
        gsMsg = "沒有此貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmLstCurr.SetFocus
       Exit Function
    End If
    
    Chk_cboItmLstCurr = True
End Function

Private Function chk_txtItmLstExcr() As Boolean
    chk_txtItmLstExcr = False
    
    If Trim(txtItmLstExcr.Text) = "" Or Trim(To_Value(txtItmLstExcr.Text)) = 0 Then
        gsMsg = "必需輸入對換率!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtItmLstExcr.SetFocus
        Exit Function
    End If
    
    If To_Value(txtItmLstExcr.Text) > 9999.999999 Then
        gsMsg = "對換率超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtItmLstExcr.SetFocus
        Exit Function
    End If
    txtItmLstExcr.Text = Format(txtItmLstExcr.Text, gsExrFmt)
    
    chk_txtItmLstExcr = True
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

