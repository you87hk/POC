VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIC001 
   Caption         =   "訂貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmIC001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11880
      OleObjectBlob   =   "frmIC001.frx":030A
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   800
      Width           =   1935
   End
   Begin VB.ComboBox cboSaleCode 
      Height          =   300
      Left            =   5280
      TabIndex        =   4
      Top             =   1150
      Width           =   1455
   End
   Begin VB.ComboBox cboRefDocNo 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1150
      Width           =   1935
   End
   Begin VB.Frame fraKey 
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtRevNo 
         Height          =   324
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "12345678901234567890"
         Top             =   300
         Width           =   408
      End
      Begin MSMask.MaskEdBox medDocDate 
         Height          =   285
         Left            =   9840
         TabIndex        =   2
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDocDate 
         Caption         =   "DOCDATE"
         Height          =   255
         Left            =   7845
         TabIndex        =   14
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label lblRevNo 
         Caption         =   "REVNO"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   1215
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
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSaleCode 
         Caption         =   "SALECODE"
         Height          =   240
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lblDspSaleDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6600
         TabIndex        =   10
         Top             =   660
         Width           =   4575
      End
      Begin VB.Label lblRefDocNo 
         Caption         =   "REFDOCNO"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIC001.frx":66AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            Key             =   "Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6675
      Left            =   120
      OleObjectBlob   =   "frmIC001.frx":69CD
      TabIndex        =   5
      Top             =   1800
      Width           =   11535
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
Attribute VB_Name = "frmIC001"
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
Private wgsTitle As String
Private wsOldRefDocNo As String
Private wsRefSrcCode As String
Private wsSrcCode As String
Private wsTrnType As String

Private Const LINENO = 0
Private Const BOOKCODE = 1
Private Const BARCODE = 2
Private Const WhsCode = 3
Private Const LOTNO = 4
Private Const BOOKNAME = 5
Private Const PUBLISHER = 6
Private Const Qty = 7
Private Const BOOKID = 8
Private Const SOID = 9

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

Private wiOpenDoc As Integer
Private wiAction As Integer
Private wiRevNo As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlRefDocID As Long
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "icStHd"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, SOID
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
      
    wsOldRefDocNo = ""

    wlKey = 0
    wlSaleID = 0
    wlRefDocID = 0
    wlLineNo = 1
    
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False

    
    Me.Caption = wsFormCaption
    
    FocusMe cboDocNo
End Sub

Private Sub cboDocNo_GotFocus()
    FocusMe cboDocNo
End Sub

Private Sub cboDocNo_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSql = "SELECT STHDDOCNO, SALENAME, STHDDOCDATE "
    wsSql = wsSql & " FROM ICSTHD, MSTSALESMAN "
    wsSql = wsSql & " WHERE STHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND STHDSALEID = SALEID "
    wsSql = wsSql & " AND STHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY STHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

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
        Call SetFieldStatus("AfrKey")
        cboRefDocNo.SetFocus
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        txtRevNo.Enabled = True
        
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKey")
        cboSaleCode.SetFocus
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
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
        Case vbKeyPageUp
            KeyCode = 0
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
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    wsSql = "SELECT STHDDOCID, STHDDOCNO, STHDREFDOCID, "
    wsSql = wsSql & "STHDDOCDATE, STHDREVNO, STDTDOCLINE, STHDSALEID, "
    wsSql = wsSql & "STDTITEMID, ITMCODE, STDTWHSCODE, STDTLOTNO, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, STDTQTY, "
    wsSql = wsSql & "STDTSOID "
    wsSql = wsSql & "FROM  icStHd, icStDT, MstSalesman, MstItem "
    wsSql = wsSql & "WHERE STHDDOCNO = '" & cboDocNo & "' "
    wsSql = wsSql & "AND STHDDOCID = STDTDOCID "
    wsSql = wsSql & "AND STHDSALEID = SALEID "
    wsSql = wsSql & "AND STDTITEMID = ITMID "
    wsSql = wsSql & "ORDER BY STDTDOCLINE "
    
    rsInvoice.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsInvoice, "STHDDOCID")
    txtRevNo.Text = Format(ReadRs(rsInvoice, "STHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsInvoice, "STHDREVNO"))
    medDocDate.Text = ReadRs(rsInvoice, "STHDDOCDATE")
    wlRefDocID = ReadRs(rsInvoice, "STHDREFDOCID")
    
    wlSaleID = To_Value(ReadRs(rsInvoice, "STHDSALEID"))
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
   Select Case wsRefSrcCode
   Case "SP"
    cboRefDocNo.Text = Get_TableInfo("soaSPHD", "SPHDDOCID =" & wlRefDocID, "SPHDDOCNO")
   Case "SR"
    cboRefDocNo.Text = Get_TableInfo("soaSRHD", "SRHDDOCID =" & wlRefDocID, "SRHDDOCNO")
   Case "PV"
    cboRefDocNo.Text = Get_TableInfo("popPVHD", "PVHDDOCID =" & wlRefDocID, "PVHDDOCNO")
   Case "PR"
    cboRefDocNo.Text = Get_TableInfo("popPRHD", "PRHDDOCID =" & wlRefDocID, "PRHDDOCNO")
   End Select
  
    
    rsInvoice.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, SOID
         Do While Not rsInvoice.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsInvoice, "STDTDOCLINE")
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsInvoice, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsInvoice, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsInvoice, "ITNAME")
             waResult(.UpperBound(1), WhsCode) = ReadRs(rsInvoice, "STDTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsInvoice, "STDTLOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsInvoice, "ITMPUBLISHER")
             waResult(.UpperBound(1), Qty) = Format(ReadRs(rsInvoice, "STDTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsInvoice, "STDTITEMID")
             waResult(.UpperBound(1), SOID) = ReadRs(rsInvoice, "STDTSOID")
             
             rsInvoice.MoveNext
         Loop
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
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
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRefDocNo.Caption = Get_Caption(waScrItm, "REFNO")
    lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(PUBLISHER).Caption = Get_Caption(waScrItm, "PUBLISHER")
        .Columns(Qty).Caption = Get_Caption(waScrItm, "QTY")
    End With
    
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
    
    wsActNam(1) = Get_Caption(waScrItm, "STADD")
    wsActNam(2) = Get_Caption(waScrItm, "STEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "STDELETE")
    wgsTitle = Get_Caption(waScrItm, "TITLE")
    
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
    Set frmIC001 = Nothing

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
            If cboRefDocNo.Enabled = True Then
                cboRefDocNo.SetFocus
            Else
                cboSaleCode.SetFocus
            End If
        End If
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
    
    Chk_medDocDate = True

End Function

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

Private Function Chk_KeyExist() As Boolean
    
    Dim rsicStHD As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT STHDSTATUS FROM icStHd WHERE STHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rsicStHD.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rsicStHD.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rsicStHD.Close
    Set rsicStHD = Nothing

End Function

Private Function Chk_KeyFld() As Boolean
        
    Chk_KeyFld = False
    
    If Chk_cboRefDocNo = False Then
        Exit Function
    End If
    
    If Chk_medDocDate = False Then
        Exit Function
    End If
    
    If Chk_cboSaleCode = False Then
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
        
    adcmdSave.CommandText = "USP_IC001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlRefDocID)
    Call SetSPPara(adcmdSave, 6, wsRefSrcCode)
    Call SetSPPara(adcmdSave, 7, Set_MedDate(medDocDate))
    Call SetSPPara(adcmdSave, 8, txtRevNo)
    Call SetSPPara(adcmdSave, 9, wsCtlPrd)
    Call SetSPPara(adcmdSave, 10, wlSaleID)
    
    Call SetSPPara(adcmdSave, 11, wsSrcCode)
    
    Call SetSPPara(adcmdSave, 12, "")
    
    For i = 1 To 10
        Call SetSPPara(adcmdSave, 13 + i - 1, "")
    Next
    
    Call SetSPPara(adcmdSave, 23, wsFormID)
    
    Call SetSPPara(adcmdSave, 24, gsUserID)
    Call SetSPPara(adcmdSave, 25, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 26)
    wsDocNo = GetSPPara(adcmdSave, 27)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_IC001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SOID))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 5, wiCtr + 1)
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, Qty))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, WhsCode))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, LOTNO))
                Call SetSPPara(adcmdSave, 9, wsTrnType)
                Call SetSPPara(adcmdSave, 10, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 11, gsUserID)
                Call SetSPPara(adcmdSave, 12, wsGenDte)
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
    If Not Chk_cboRefDocNo Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
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
                
                If Chk_NoDup2(wlCtr, waResult(wlCtr, LINENO), waResult(wlCtr, BOOKCODE), waResult(wlCtr, WhsCode), waResult(wlCtr, LOTNO)) = False Then
                    tblDetail.Row = wlCtr - 1
                    tblDetail.Col = BOOKCODE
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
            tblDetail.Col = BOOKCODE
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

    Dim newForm As New frmIC001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmIC001
    
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
 '   wsFormID = "IC0011"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
 '   wsTrnCd = "IO"
 '   wsRefSrcCode = "SP"
 '   wsSrcCode = "SOP"
 '   wsTrnType = "STKOUT"
End Sub

Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("Default")
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
           If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, gsTitle) = vbYes Then
                Call cmdCancel
           End If
           Else
                Call cmdCancel
           End If
        Case tcRefresh
            Call cmdRefresh
        Case tcPrint
            Call cmdPrint
        Case tcExit
            Unload Me
    End Select
    
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

Private Sub cboRefDocNo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboRefDocNo
    
    Select Case wsRefSrcCode
    
    Case "SP"
    wsSql = "SELECT SPHDDOCNO, SPHDDOCDATE FROM soaSPHD "
    wsSql = wsSql & " WHERE SPHDSTATUS = '1' "
    wsSql = wsSql & " AND SPHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY SPHDDOCNO "
    Case "SR"
    wsSql = "SELECT SRHDDOCNO, SRHDDOCDATE FROM soaSRHD "
    wsSql = wsSql & " WHERE SRHDSTATUS = '1' "
    wsSql = wsSql & " AND SRHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY SRHDDOCNO "
    
    Case "PV"
    wsSql = "SELECT PVHDDOCNO, PVHDDOCDATE FROM popPVHD "
    wsSql = wsSql & " WHERE PVHDSTATUS = '1' "
    wsSql = wsSql & " AND PVHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY PVHDDOCNO "
    
    Case "PR"
    wsSql = "SELECT PRHDDOCNO, PRHDDOCDATE FROM popPRHD "
    wsSql = wsSql & " WHERE PRHDSTATUS = '1' "
    wsSql = wsSql & " AND PRHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY PRHDDOCNO "
    
    End Select
    
    Call Ini_Combo(2, wsSql, cboRefDocNo.Left, cboRefDocNo.Top + cboRefDocNo.Height, tblCommon, wsFormID, "TBLREFNO", Me.Width, Me.Height)
            
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
        
        If wiAction = AddRec And wsOldRefDocNo <> cboRefDocNo.Text Then Call Get_RefDoc
        
        cboSaleCode.SetFocus
    End If
    
End Sub

Private Function Chk_cboRefDocNo() As Boolean
    
Dim wsStatus As String
    
    Chk_cboRefDocNo = False
    
    If Trim(cboRefDocNo.Text) = "" Then
        gsMsg = "未有輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
        
    If Chk_TrnHdDocNo(wsRefSrcCode, cboRefDocNo, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        
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
        
    End If
    
    Chk_cboRefDocNo = True

End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstCUSTOMER "
    wsSql = wsSql & "WHERE CUSID = " & wlCusID
    rsDefVal.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        wlSaleID = ReadRs(rsDefVal, "CUSSALEID")
    Else
        wlSaleID = 0
    End If
    rsDefVal.Close
    Set rsDefVal = Nothing
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
End Sub

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = LINENO To SOID
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
                Case WhsCode
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case LOTNO
                    .Columns(wiCtr).Width = 1000
                    '.Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Visible = False
                Case BOOKNAME
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case PUBLISHER
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
                Case Qty
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case BOOKID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case SOID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
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
    Dim wsSoID As String
    
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
                
                If Chk_grdBookCode(cboRefDocNo, .Columns(ColIndex).Text, wsBookID, wsBookCode, wsBarCode, wsBookName, wsPub, wdPrice, wdDisPer, wsWhsCode, wsLotNo, wdQty, wsSoID) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(LINENO).Text = wlLineNo
                .Columns(BOOKID).Text = wsBookID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(PUBLISHER).Text = wsPub
                .Columns(WhsCode).Text = wsWhsCode
                .Columns(LOTNO).Text = wsLotNo
                .Columns(SOID).Text = wsSoID
                .Columns(Qty).Text = ""
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(ColIndex).Text) <> wsBookCode Then
                    .Columns(ColIndex).Text = wsBookCode
                End If
        
             Case WhsCode
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
             Case LOTNO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdLotNo(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If

            Case Qty
            
                If ColIndex = Qty Then
                    If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
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
    
    On Error GoTo tblDetail_ButtonClick_Err
    
    With tblDetail
        Select Case ColIndex
            Case BOOKCODE
                
                Select Case wsRefSrcCode
    
                Case "SP"
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, SOASPHD, SOASPDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SPHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND SPHDDOCID = SPDTDOCID "
                    wsSql = wsSql & " AND SPDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, SOASPHD, SOASPDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SPHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND SPHDDOCID = SPDTDOCID "
                    wsSql = wsSql & " AND SPDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                Case "SR"
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, SOASRHD, SOASRDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SRHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND SRHDDOCID = SRDTDOCID "
                    wsSql = wsSql & " AND SRDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, SOASRHD, SOASRDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SRHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND SRHDDOCID = SRDTDOCID "
                    wsSql = wsSql & " AND SRDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Case "PV"
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, POPPVHD, POPPVDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND PVHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND PVHDDOCID = PVDTDOCID "
                    wsSql = wsSql & " AND PVDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, POPPVHD, POPPVDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND PVHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND PVHDDOCID = PVDTDOCID "
                    wsSql = wsSql & " AND PVDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Case "PR"
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, POPPRHD, POPPRDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND PRHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND PRHDDOCID = PRDTDOCID "
                    wsSql = wsSql & " AND PRDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM "
                    wsSql = wsSql & " FROM mstITEM, POPPRHD, POPPRDT "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND PRHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
                    wsSql = wsSql & " AND PRHDDOCID = PRDTDOCID "
                    wsSql = wsSql & " AND PRDTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                End Select
                
                Call Ini_Combo(4, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLBOOKCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case WhsCode
                
                wsSql = "SELECT WHSCODE, WHSDESC FROM mstWareHouse "
                wsSql = wsSql & " WHERE WHSSTATUS <> '2' AND WHSCODE LIKE '%" & Set_Quote(.Columns(WhsCode).Text) & "%' "
                wsSql = wsSql & " ORDER BY WHSCODE "
                
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
           End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDetail ButtonClick!"
 
    
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
                Case LINENO
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case BOOKCODE
                    KeyCode = vbDefault
                    .Col = Qty
                Case Qty
                    KeyCode = vbKeyDown
                    .Col = BOOKCODE
            End Select
        Case vbKeyLeft
               KeyCode = vbDefault
            Select Case .Col
                Case Qty
                    .Col = BOOKCODE
                Case BOOKCODE
                    .Col = Qty
                Case LINENO
                    .Col = .Col + 1
            End Select
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case BOOKCODE
                    .Col = Qty
                Case LINENO
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
        
        Case Qty
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
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(cboRefDocNo, .Columns(BOOKCODE).Text, "", "", "", "", "", 0, 0, "", "", 0, "")
                Case WhsCode
                    Call Chk_grdWhsCode(.Columns(WhsCode).Text)
                 Case LOTNO
                    Call Chk_grdLotNo(.Columns(LOTNO).Text)
                Case Qty
                    Call Chk_grdQty(.Columns(Qty).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdBookCode(inSoNo As String, inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double, outWhsCode As String, outLotNo As String, outQty As Double, outSOID As String) As Boolean
    
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入書號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
        Exit Function
    End If
    
    Select Case wsRefSrcCode
    Case "SP"
    
    wsSql = "SELECT SPDTDOCID DOCID , SPDTITEMID ITMID, ITMCODE, SPDTITEMDESC ITNAME, ITMBARCODE, ITMPUBLISHER, SPDTWHSCODE WHSCODE , SPDTLOTNO LOTNO, SPDTSOID SOID "
    wsSql = wsSql & " FROM mstITEM, soaSPHD, soaSPDT "
    wsSql = wsSql & " WHERE SPHDDOCID = SPDTDOCID "
    wsSql = wsSql & " AND SPDTITEMID = ITMID "
    wsSql = wsSql & " AND SPHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
    
    Case "SR"
    
    wsSql = "SELECT SRDTDOCID DOCID , SRDTITEMID ITMID, ITMCODE, SRDTITEMDESC ITNAME, ITMBARCODE, ITMPUBLISHER, SRDTWHSCODE WHSCODE , SRDTLOTNO LOTNO, SRDTSOID SOID "
    wsSql = wsSql & " FROM mstITEM, soaSRHD, soaSRDT "
    wsSql = wsSql & " WHERE SRHDDOCID = SRDTDOCID "
    wsSql = wsSql & " AND SRDTITEMID = ITMID "
    wsSql = wsSql & " AND SRHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
    
    Case "PV"
    
    wsSql = "SELECT PVDTDOCID DOCID , PVDTITEMID ITMID, ITMCODE, PVDTITEMDESC ITNAME, ITMBARCODE, ITMPUBLISHER, PVDTWHSCODE WHSCODE , PVDTLOTNO LOTNO, PVDTPOID SOID "
    wsSql = wsSql & " FROM mstITEM, popPVHD, popPVDT "
    wsSql = wsSql & " WHERE PVHDDOCID = PVDTDOCID "
    wsSql = wsSql & " AND PVDTITEMID = ITMID "
    wsSql = wsSql & " AND PVHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
   
    Case "PR"
    
    wsSql = "SELECT PRDTDOCID DOCID , PRDTITEMID ITMID, ITMCODE, PRDTITEMDESC ITNAME, ITMBARCODE, ITMPUBLISHER, PRDTWHSCODE WHSCODE , PRDTLOTNO LOTNO , PRDTPOID SOID "
    wsSql = wsSql & " FROM mstITEM, popPRHD, popPRDT "
    wsSql = wsSql & " WHERE PRHDDOCID = PRDTDOCID "
    wsSql = wsSql & " AND PRDTITEMID = ITMID "
    wsSql = wsSql & " AND PRHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
   
    End Select
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       outAccID = ReadRs(rsDes, "ITMID")
       outAccNo = ReadRs(rsDes, "ITMCODE")
       OutName = ReadRs(rsDes, "ITNAME")
       OutBarCode = ReadRs(rsDes, "ITMBARCODE")
       outPub = ReadRs(rsDes, "ITMPUBLISHER")
       outWhsCode = ReadRs(rsDes, "WHSCODE")
       outLotNo = ReadRs(rsDes, "LOTNO")
       'outQty = Get_SoBalQty(wsTrnCd, wlKey, ReadRs(rsDes, "SODTDOCID"), outAccID, outWhsCode, outLotNo)
       outQty = Format(0, gsAmtFmt)
       outSOID = To_Value(ReadRs(rsDes, "SOID"))
       
       Chk_grdBookCode = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        outPub = ""
        outLotNo = ""
        outWhsCode = ""
        outQty = Format(0, gsAmtFmt)
        outSOID = "0"
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdOrdItm(inSoNo As String, inItmNo As String, inWhsCode As String, InLotNo As String) As Boolean
    
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset

    
    Select Case wsRefSrcCode
    
    Case "SP"
    wsSql = "SELECT SPDTITEMID "
    wsSql = wsSql & " FROM mstITEM, soaSpHd, soaSpDt "
    wsSql = wsSql & " WHERE SPHDDOCID = SPDTDOCID "
    wsSql = wsSql & " AND SPDTITEMID = ITMID "
    wsSql = wsSql & " AND SPHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSql = wsSql & " AND SPDTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSql = wsSql & " AND SPDTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSql = wsSql & " AND SPHDSTATUS NOT IN ('2' , '3')"
    
    Case "SR"
    wsSql = "SELECT SRDTITEMID "
    wsSql = wsSql & " FROM mstITEM, soaSRHD, soaSRDT "
    wsSql = wsSql & " WHERE SRHDDOCID = SRDTDOCID "
    wsSql = wsSql & " AND SRDTITEMID = ITMID "
    wsSql = wsSql & " AND SRHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSql = wsSql & " AND SRDTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSql = wsSql & " AND SRDTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSql = wsSql & " AND SRHDSTATUS NOT IN ('2' , '3')"
    
    Case "PV"
    wsSql = "SELECT PVDTITEMID "
    wsSql = wsSql & " FROM mstITEM, popPVHD, popPVDT "
    wsSql = wsSql & " WHERE PVHDDOCID = PVDTDOCID "
    wsSql = wsSql & " AND PVDTITEMID = ITMID "
    wsSql = wsSql & " AND PVHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSql = wsSql & " AND PVDTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSql = wsSql & " AND PVDTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSql = wsSql & " AND PVHDSTATUS NOT IN ('2' , '3')"
    
    Case "PR"
    wsSql = "SELECT PRDTITEMID "
    wsSql = wsSql & " FROM mstITEM, popPRHD, popPRDT "
    wsSql = wsSql & " WHERE PRHDDOCID = PRDTDOCID "
    wsSql = wsSql & " AND PRDTITEMID = ITMID "
    wsSql = wsSql & " AND PRHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSql = wsSql & " AND PRDTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSql = wsSql & " AND PRDTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSql = wsSql & " AND PRHDSTATUS NOT IN ('2' , '3')"
        
    
    End Select
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       Chk_grdOrdItm = True
    Else
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdOrdItm = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function



Private Function Chk_grdWhsCode(inNo As String) As Boolean
    
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdWhsCode = False
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入貨倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    wsSql = "SELECT *  FROM mstWareHouse"
    wsSql = wsSql & " WHERE WHSCODE = '" & Set_Quote(inNo) & "' "
    wsSql = wsSql & " AND WHSSTATUS = '1' "
       
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    
    Dim wsSql As String
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
                   Trim(waResult(inRow, PUBLISHER)) = "" And _
                   Trim(waResult(inRow, Qty)) = "" And _
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
        
        If Chk_grdBookCode(cboRefDocNo, waResult(LastRow, BOOKCODE), "", "", "", "", "", 0, 0, "", "", 0, "") = False Then
            .Col = BOOKCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdWhsCode(waResult(LastRow, WhsCode)) = False Then
                .Col = WhsCode
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdLotNo(waResult(LastRow, LOTNO)) = False Then
                .Col = LOTNO
                .Row = LastRow
                Exit Function
        End If
        
        
        If Chk_grdQty(waResult(LastRow, Qty)) = False Then
                .Col = Qty
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdOrdItm(cboRefDocNo, waResult(LastRow, BOOKCODE), waResult(LastRow, WhsCode), waResult(LastRow, LOTNO)) = False Then
            .Col = BOOKCODE
            .Row = LastRow
            Exit Function
        End If
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
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
        
    adcmdDelete.CommandText = "USP_IC001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlRefDocID)
    Call SetSPPara(adcmdDelete, 6, wsRefSrcCode)
    Call SetSPPara(adcmdDelete, 7, Set_MedDate(medDocDate))
    Call SetSPPara(adcmdDelete, 8, txtRevNo)
    Call SetSPPara(adcmdDelete, 9, "")
    Call SetSPPara(adcmdDelete, 10, wlSaleID)
    
    Call SetSPPara(adcmdDelete, 11, wsSrcCode)
    
    Call SetSPPara(adcmdDelete, 12, "")
    
    For i = 1 To 10
        Call SetSPPara(adcmdDelete, 13 + i - 1, "")
    Next
    
    Call SetSPPara(adcmdDelete, 23, wsFormID)
    
    Call SetSPPara(adcmdDelete, 24, gsUserID)
    Call SetSPPara(adcmdDelete, 25, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 26)
    wsDocNo = GetSPPara(adcmdDelete, 27)
    
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
        
        gsMsg = "你是否確定不儲存現時之變更而離開?"
        If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbOK Then
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

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.cboRefDocNo.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            
            Me.cboSaleCode.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
            Me.cboRefDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
            
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            If wiAction = AddRec Then
                Me.cboRefDocNo.Enabled = True
            Else
                Me.cboRefDocNo.Enabled = False
            End If
            
            
            Me.txtRevNo.Enabled = True
            Me.medDocDate.Enabled = True
            
            Me.cboSaleCode.Enabled = True
            
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
        .TableKey = "STHDDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSql As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "Doc No."
    vFilterAry(1, 2) = "STHDDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "STHDDocDate"
    
    vFilterAry(3, 1) = "Salesman #"
    vFilterAry(3, 2) = "SaleCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "StHDDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "STHDDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Salesman #"
    vAry(3, 2) = "SaleCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Sales Name"
    vAry(4, 2) = "SalesName"
    vAry(4, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSql = "SELECT icStHd.STHDDocNo, icStHd.STHDDocDate, MstMstSalesman.SaleCode,  MstSalesman.SaleName "
        wsSql = wsSql + "FROM MstSalesman, icStHd "
        .sBindSQL = wsSql
        .sBindWhereSQL = "WHERE icStHd.STHDStatus = '1' And icStHd.STHDSaleID = MstSalesman.SaleID "
        .sBindOrderSQL = "ORDER BY icStHd.STHDDocNo"
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
    
End Sub

Private Sub cbosaleCode_GotFocus()
    FocusMe cboSaleCode
End Sub

Private Sub cboSaleCode_LostFocus()
    FocusMe cboSaleCode, True
End Sub

Private Sub cbosaleCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_KeyFld = False Then
            Exit Sub
        End If
        
        If tblDetail.Enabled = True Then
            tblDetail.Col = BOOKCODE
            tblDetail.SetFocus
        Else
            medDocDate.SetFocus
        End If
    End If
End Sub

Private Sub cbosaleCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode
    
    wsSql = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
    wsSql = wsSql & "AND SaleStatus = '1' "
    wsSql = wsSql & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSql, cboSaleCode.Left, cboSaleCode.Top + cboSaleCode.Height, tblCommon, wsFormID, "TBLSALECOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboSaleCode() As Boolean
Dim wsDesc As String

    Chk_cboSaleCode = False
     
    If Trim(cboSaleCode.Text) = "" Then
        gsMsg = "必需輸入倉務員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        Exit Function
    End If
    
    If Chk_Salesman(cboSaleCode, wlSaleID, wsDesc) = False Then
        gsMsg = "沒有此倉務員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        lblDspSaleDesc = ""
       Exit Function
    End If
    
    lblDspSaleDesc = wsDesc
    
    Chk_cboSaleCode = True
    
End Function

Private Sub txtRevNo_LostFocus()
    FocusMe txtRevNo, True
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Dim wsCurRecLn2 As String
    Dim wsCurRecLn3 As String
    
    Chk_NoDup = False
    
    wsCurRecLn = tblDetail.Columns(BOOKCODE)
    'wsCurRecLn2 = tblDetail.Columns(WhsCode)
    'wsCurRecLn3 = tblDetail.Columns(LOTNO)
   
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           'If wsCurRecLn = waResult(wlCtr, BOOKCODE) And _
              wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
              wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
            If wsCurRecLn = waResult(wlCtr, BOOKCODE) Then
              gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
              MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
              Exit Function
           End If
        End If
    Next
    
    Chk_NoDup = True

End Function

Private Function Chk_NoDup2(ByRef inRow As Long, ByVal wsCurRec As String, ByVal wsCurRecLn As String, ByVal wsCurRecLn2 As String, ByVal wsCurRecLn3 As String) As Boolean
    
    Dim wlCtr As Long
     
    Chk_NoDup2 = False
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           'If wsCurRec = waResult(wlCtr, LINENO) And _
              wsCurRecLn = waResult(wlCtr, BOOKCODE) And _
              wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
              wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
            If wsCurRec = waResult(wlCtr, BOOKCODE) Then
                gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                inRow = To_Value(waResult(wlCtr, LINENO))
                Exit Function
           End If
        End If
    Next
    
    Chk_NoDup2 = True

End Function

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
    
    '' form delcare
    'Private waPopUpSub As New XArrayDB
    
    '' form unload
    'Set waPopUpSub = Nothing
    
    ''   addin ini_caption
    '    Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
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

Public Sub SetButtonStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
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
            
            End With
    End Select
End Sub

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    'If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTINV002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSql = wsSql & "'" & "" & "', "
    wsSql = wsSql & "'" & String(10, "z") & "', "
    wsSql = wsSql & "'" & "0000/00/00" & "', "
    wsSql = wsSql & "'" & "9999/99/99" & "', "
    wsSql = wsSql & "'" & "%" & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then wsRptName = "C" + "RPTINV002"
    
    NewfrmPrint.ReportID = "INV002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "INV002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdRefresh()
    Dim wiCtr As Integer
    Dim wsBookID As String
    Dim wdDisPer As Double
    Dim wdNewDisPer As Double

    
    If waResult.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                wsBookID = waResult(wiCtr, BOOKID)
                'wdNewDisPer = Get_SaleDiscount(cboNatureCode.Text, wlCusID, wsBookID)
                'If wdDisPer <> wdNewDisPer Then
                '    waResult(wiCtr, DisPer) = Format(wdNewDisPer, gsAmtFmt)
                '    waResult(wiCtr, Dis) = Format(To_Value(waResult(wiCtr, Amt)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                '    waResult(wiCtr, Disl) = Format(To_Value(waResult(wiCtr, Amtl)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                '    waResult(wiCtr, Net) = Format(To_Value(waResult(wiCtr, Amt)) - To_Value(waResult(wiCtr, Dis)), gsAmtFmt)
                '    waResult(wiCtr, Netl) = Format(To_Value(waResult(wiCtr, Amtl)) - To_Value(waResult(wiCtr, Disl)), gsAmtFmt)
                'End If
            End If
        Next
   
        tblDetail.ReBind
        tblDetail.FirstRow = 0
    
    End If
End Sub

Private Function Get_RefDoc() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    
    Get_RefDoc = False
    Select Case wsRefSrcCode
    Case "SP"
    
    wsSql = "SELECT SPHDDOCID DOCID, SPDTITEMID ITMID, SPDTSOID SOID, "
    wsSql = wsSql & "ITMCODE, SPDTWHSCODE WHSCODE , SPDTLOTNO LOTNO, ITMBARCODE, SPDTITEMDESC ITNAME, "
    wsSql = wsSql & "ITMPUBLISHER "
    wsSql = wsSql & "FROM  soaSPHD, soaSPDT, mstITEM "
    wsSql = wsSql & "WHERE SPHDDOCNO = '" & cboRefDocNo & "' "
    wsSql = wsSql & "AND SPHDDOCID = SPDTDOCID "
    wsSql = wsSql & "AND SPDTITEMID = ITMID "
    wsSql = wsSql & "ORDER BY SPDTDOCLINE "
    
    Case "SR"
    
    wsSql = "SELECT SRHDDOCID DOCID, SRDTITEMID ITMID, SRDTSOID SOID, "
    wsSql = wsSql & "ITMCODE, SRDTWHSCODE WHSCODE , SRDTLOTNO LOTNO, ITMBARCODE, SRDTITEMDESC ITNAME, "
    wsSql = wsSql & "ITMPUBLISHER "
    wsSql = wsSql & "FROM  soaSRHD, soaSRDT, mstITEM "
    wsSql = wsSql & "WHERE SRHDDOCNO = '" & cboRefDocNo & "' "
    wsSql = wsSql & "AND SRHDDOCID = SRDTDOCID "
    wsSql = wsSql & "AND SRDTITEMID = ITMID "
    wsSql = wsSql & "ORDER BY SRDTDOCLINE "
    
    Case "PV"
    
    wsSql = "SELECT PVHDDOCID DOCID, PVDTITEMID ITMID, PVDTPOID SOID, "
    wsSql = wsSql & "ITMCODE, PVDTWHSCODE WHSCODE , PVDTLOTNO LOTNO, ITMBARCODE, PVDTITEMDESC ITNAME, "
    wsSql = wsSql & "ITMPUBLISHER "
    wsSql = wsSql & "FROM  POPPVHD, POPPVDT, mstITEM "
    wsSql = wsSql & "WHERE PVHDDOCNO = '" & cboRefDocNo & "' "
    wsSql = wsSql & "AND PVHDDOCID = PVDTDOCID "
    wsSql = wsSql & "AND PVDTITEMID = ITMID "
    wsSql = wsSql & "ORDER BY PVDTDOCLINE "
    
    Case "PR"
    
    wsSql = "SELECT PRHDDOCID DOCID, PRDTITEMID ITMID, PRDTPOID SOID, "
    wsSql = wsSql & "ITMCODE, PRDTWHSCODE WHSCODE , PRDTLOTNO LOTNO, ITMBARCODE, PRDTITEMDESC ITNAME, "
    wsSql = wsSql & "ITMPUBLISHER "
    wsSql = wsSql & "FROM  POPPRHD, POPPRDT, mstITEM "
    wsSql = wsSql & "WHERE PRHDDOCNO = '" & cboRefDocNo & "' "
    wsSql = wsSql & "AND PRHDDOCID = PRDTDOCID "
    wsSql = wsSql & "AND PRDTITEMID = ITMID "
    wsSql = wsSql & "ORDER BY PRDTDOCLINE "
    
    End Select
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsOldRefDocNo = cboRefDocNo.Text
    wlRefDocID = ReadRs(rsRcd, "DOCID")
    
    wlLineNo = 1
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, SOID
         Do While Not rsRcd.EOF
            
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITNAME")
             waResult(.UpperBound(1), WhsCode) = ReadRs(rsRcd, "WHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsRcd, "LOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsRcd, "ITMPUBLISHER")
             'waResult(.UpperBound(1), Qty) = Format(wdBalQty, gsQtyFmt)
             waResult(.UpperBound(1), Qty) = ""
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "ITMID")
             waResult(.UpperBound(1), SOID) = ReadRs(rsRcd, "SOID")
             wlLineNo = wlLineNo + 1
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Get_RefDoc = True
    
End Function

Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property
Public Property Let RefSrcCode(InRefSrcCode As String)
    wsRefSrcCode = InRefSrcCode
End Property
Public Property Let SrcCode(InSrcCode As String)
    wsSrcCode = InSrcCode
End Property
Public Property Let TrnType(InTrnType As String)
    wsTrnType = InTrnType
End Property


