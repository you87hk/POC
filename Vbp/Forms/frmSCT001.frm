VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSCT001 
   Caption         =   "訂貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmSCT001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   7200
      OleObjectBlob   =   "frmSCT001.frx":030A
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ComboBox cboWhsCodeFr 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1425
      Width           =   1935
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   675
      Width           =   1935
   End
   Begin VB.ComboBox cboSaleCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1050
      Width           =   1935
   End
   Begin VB.Frame fraKey 
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtRmk 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         Top             =   1320
         Width           =   7200
      End
      Begin MSMask.MaskEdBox medDocDate 
         Height          =   285
         Left            =   9840
         TabIndex        =   1
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblRmk 
         Caption         =   "RMK"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblDspVoucher 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1560
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblVoucher 
         Caption         =   "MLCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label lblWhsCodeFr 
         Caption         =   "WHSFR"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   975
         Width           =   1440
      End
      Begin VB.Label lblDocDate 
         Caption         =   "DOCDATE"
         Height          =   255
         Left            =   8640
         TabIndex        =   10
         Top             =   240
         Width           =   1320
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
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSaleCode 
         Caption         =   "SALECODE"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lblDspSaleDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3600
         TabIndex        =   7
         Top             =   540
         Width           =   1935
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
            Picture         =   "frmSCT001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSCT001.frx":66AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   4995
      Left            =   120
      OleObjectBlob   =   "frmSCT001.frx":69CD
      TabIndex        =   4
      Top             =   3000
      Width           =   11535
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Post"
            Object.ToolTipText     =   "Post"
            ImageIndex      =   2
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
            Object.Visible         =   0   'False
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
   Begin VB.Label lblTrnQty 
      Caption         =   "TRNQTY"
      Height          =   240
      Left            =   600
      TabIndex        =   14
      Top             =   8205
      Width           =   1500
   End
   Begin VB.Label lblTrnAmt 
      Caption         =   "TRNAMT"
      Height          =   240
      Left            =   7200
      TabIndex        =   13
      Top             =   8205
      Width           =   1500
   End
   Begin VB.Label lblDspTrnQty 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   2160
      TabIndex        =   12
      Top             =   8160
      Width           =   2265
   End
   Begin VB.Label lblDspTrnAmt 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   8760
      TabIndex        =   11
      Top             =   8160
      Width           =   2265
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
Attribute VB_Name = "frmSCT001"
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
Private wsSrcCode As String
Private wsTrnType As String

Private Const LINENO = 0
Private Const DocNo = 1
Private Const CTLPRD = 2
Private Const NOOFITM = 3
Private Const ADJQTY = 4
Private Const ADJAMT = 5
Private Const UPDATE = 6
Private Const DocID = 7

Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"

Private Const tcPost = "Post"

Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"
Private Const tcRefresh = "Refresh"
Private Const tcPrint = "Print"

Private wiOpenDoc As Integer
Private wiAction As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlLineNo As Long


Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "ICSTKCNT"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, DocID
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
    

      
    wlKey = 0
    wlSaleID = 0
    wlLineNo = 1
    wbReadOnly = False
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
End Sub

Private Sub cboDocNo_GotFocus()
    FocusMe cboDocNo
End Sub

Private Sub cboDocNo_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSQL = "SELECT SCHDDOCNO, SCHDDOCDATE "
    wsSQL = wsSQL & " FROM ICSTKCNT "
    wsSQL = wsSQL & " WHERE SCHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND SCHDSTATUS in ('1','4') "
    wsSQL = wsSQL & " AND SCHDTRNCODE  = '" & wsTrnCd & "' "
    wsSQL = wsSQL & " ORDER BY SCHDDOCNO DESC "
    
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
        medDocDate.Text = Dsp_Date(Now)
    
        Call SetButtonStatus("AfrKeyAdd")
        Call SetFieldStatus("AfrKey")
        cboSaleCode.SetFocus
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKey")
        cboSaleCode.SetFocus
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
End Sub



Private Sub cboWhsCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboWhsCodeFr
    
    wsSQL = "SELECT WhsCode , WhsDesc FROM MstWareHouse WHERE WhsStatus = '1' "
    wsSQL = wsSQL & " AND WhsCode LIKE '%" & IIf(cboWhsCodeFr.SelLength > 0, "", Set_Quote(cboWhsCodeFr.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY WhsCode "
    Call Ini_Combo(2, wsSQL, cboWhsCodeFr.Left, cboWhsCodeFr.Top + cboWhsCodeFr.Height, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWhsCodeFr_GotFocus()
    FocusMe cboWhsCodeFr
End Sub

Private Sub cboWhsCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCodeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboWhsCodeFr() = False Then
            Exit Sub
        End If
        
        txtRmk.SetFocus
        
        
    End If
End Sub

Private Sub cboWhsCodeFr_LostFocus()
    FocusMe cboWhsCodeFr, True
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
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    wsSQL = "SELECT SCHDDOCID, SCHDDOCNO, SCHDDOCDATE, SCHDSTAFFID, SCHDREMARK, SALECODE, SALENAME, SCHDTRNCODE, SCHDWHSCODE, SCHDVOUCNO "
    wsSQL = wsSQL & "FROM ICSTKCNT, MSTSALESMAN "
    wsSQL = wsSQL & "WHERE SCHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
    wsSQL = wsSQL & "AND SCHDSTAFFID *= SALEID "
    
    
    rsInvoice.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsInvoice, "SCHDDOCID")
    medDocDate.Text = ReadRs(rsInvoice, "SCHDDOCDATE")
    
    wlSaleID = To_Value(ReadRs(rsInvoice, "SCHDSTAFFID"))
    
    cboSaleCode.Text = ReadRs(rsInvoice, "SALECODE")
    lblDspSaleDesc = ReadRs(rsInvoice, "SALENAME")
    
    cboWhsCodeFr = ReadRs(rsInvoice, "SCHDWHSCODE")
    lblDspVoucher.Caption = ReadRs(rsInvoice, "SCHDVOUCNO")
    txtRmk.Text = ReadRs(rsInvoice, "SCHDREMARK")
       
    rsInvoice.Close
    Set rsInvoice = Nothing
       
    
    wsSQL = "SELECT SJHDDOCID, SJHDDOCNO, SJHDCTLPRD, SJHDSTATUS, COUNT(SJDTID) NOOFITEM, SUM(SJDTTRNQTY) QTY, SUM(SJDTTRNAMT) AMT "
    wsSQL = wsSQL & "FROM ICSTKADJ, ICSTKADJDT "
    wsSQL = wsSQL & "WHERE SJHDSCID = " & wlKey & " "
    wsSQL = wsSQL & "AND SJHDTRNCODE  = 'SK' "
    wsSQL = wsSQL & "AND SJHDDOCID  = SJDTDOCID "
    wsSQL = wsSQL & "AND SJHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & "GROUP BY SJHDDOCID, SJHDDOCNO, SJHDCTLPRD, SJHDSTATUS "
    
    
    rsInvoice.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount > 0 Then
    rsInvoice.MoveFirst
    
    With waResult
        .ReDim 0, -1, LINENO, DocID
        Do While Not rsInvoice.EOF
            wiCtr = wiCtr + 1
            .AppendRows
            waResult(.UpperBound(1), LINENO) = wiCtr
            waResult(.UpperBound(1), DocNo) = ReadRs(rsInvoice, "SJHDDOCNO")
            waResult(.UpperBound(1), CTLPRD) = ReadRs(rsInvoice, "SJHDCTLPRD")
            waResult(.UpperBound(1), NOOFITM) = Format(ReadRs(rsInvoice, "NOOFITEM"), gsQtyFmt)
            waResult(.UpperBound(1), ADJQTY) = Format(ReadRs(rsInvoice, "QTY"), gsQtyFmt)
            waResult(.UpperBound(1), ADJAMT) = Format(ReadRs(rsInvoice, "AMT"), gsAmtFmt)
            waResult(.UpperBound(1), UPDATE) = IIf(ReadRs(rsInvoice, "SJHDSTATUS") = "1", "N", "Y")
            waResult(.UpperBound(1), DocID) = ReadRs(rsInvoice, "SJHDDOCID")
            rsInvoice.MoveNext
        Loop
        wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    
    End If
    
    rsInvoice.Close
    Set rsInvoice = Nothing
    Call Calc_Total
    
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblVoucher.Caption = Get_Caption(waScrItm, "VOUCHNO")
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblRmk.Caption = Get_Caption(waScrItm, "REMARK")
    
    lblTrnAmt.Caption = Get_Caption(waScrItm, "TRNAMT")
    lblTrnQty.Caption = Get_Caption(waScrItm, "TRNQTY")
    

    
    'FRADOCTYPE.Caption = Get_Caption(waScrItm, "DOCTYPE")
    lblWhsCodeFr.Caption = Get_Caption(waScrItm, "WHSCODEFR")
    
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "SLINENO")
        .Columns(DocNo).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(CTLPRD).Caption = Get_Caption(waScrItm, "SCTLPRD")
        .Columns(NOOFITM).Caption = Get_Caption(waScrItm, "SNOOFITM")
        .Columns(ADJQTY).Caption = Get_Caption(waScrItm, "SADJQTY")
        .Columns(ADJAMT).Caption = Get_Caption(waScrItm, "SADJAMT")
        .Columns(UPDATE).Caption = Get_Caption(waScrItm, "SUPDATE")
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
    
    wsActNam(1) = Get_Caption(waScrItm, "ADJADD")
    wsActNam(2) = Get_Caption(waScrItm, "ADJEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "ADJDELETE")
    
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
    Set frmSCT001 = Nothing

End Sub

Private Sub medDocDate_GotFocus()
    
  FocusMe medDocDate
    
End Sub

Private Sub medDocDate_LostFocus()

    FocusMe medDocDate, True
    
End Sub

Private Sub medDocDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        If Chk_medDocDate = False Then
            Exit Sub
        End If
            
        cboSaleCode.SetFocus
        
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
        wcCombo.Text = tblCommon.Columns(0).Text
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
    
    Dim rsicStHD As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT SCHDSTATUS FROM ICSTKCNT WHERE SCHDDOCNO = '" & Set_Quote(cboDocNo) & "' AND SCHDTRNCODE  = '" & wsTrnCd & "' "
    rsicStHD.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
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
    
 '   If Chk_cboSaleCode = False Then
 '       Exit Function
 '   End If
    
    If Chk_cboWhsCodeFr = False Then
        Exit Function
    End If
    
    If Chk_medDocDate = False Then
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
    Dim wsTrnType As String
    Dim wdQty As Double
    
    
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
        
    adcmdSave.CommandText = "USP_SCT001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, Set_MedDate(medDocDate))
    Call SetSPPara(adcmdSave, 6, wsCtlPrd)
    Call SetSPPara(adcmdSave, 7, Trim(cboWhsCodeFr.Text))
    Call SetSPPara(adcmdSave, 8, wlSaleID)
    
    Call SetSPPara(adcmdSave, 9, wsSrcCode)
    Call SetSPPara(adcmdSave, 10, txtRmk.Text)
    Call SetSPPara(adcmdSave, 11, wsFormID)
    
    Call SetSPPara(adcmdSave, 12, gsUserID)
    Call SetSPPara(adcmdSave, 13, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 14)
    wsDocNo = GetSPPara(adcmdSave, 15)
    
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
    
    If Not Chk_medDocDate Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    
    If Not Chk_cboWhsCodeFr Then Exit Function
    
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cmdNew()

    Dim newForm As New frmSCT001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmSCT001
    
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
    wsFormID = "SCT001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "SC"
    wsSrcCode = "ICP"
    
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



Private Sub tblDetail_Click()

        If Chk_KeyFld = False Then
            Exit Sub
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
        Case tcPost
            Call cmdPost
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
            Call cmdPrint
        Case tcExit
            Unload Me
    End Select
    
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
        
        
        For wiCtr = LINENO To DocID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case LINENO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Locked = True
                Case DocNo
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Button = True
                Case CTLPRD
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                Case UPDATE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 1
                Case NOOFITM
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case ADJQTY
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case ADJAMT
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                   ' .Columns(wiCtr).Visible = False
                Case DocID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
       ' .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .UPDATE
    End With

End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    
    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    
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
        Case DocNo
            
            
        
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

       Case vbKeyReturn
            Select Case .Col
               Case UPDATE
                    KeyCode = vbKeyDown
                    .Col = LINENO
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
        Case vbKeyLeft
               KeyCode = vbDefault
            Select Case .Col
                Case LINENO
                    .Col = UPDATE
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col - 1
            End Select
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
               Case UPDATE
                    KeyCode = vbKeyDown
                    .Col = LINENO
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
            
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = LINENO
        End If
        
        Call Calc_Total
        
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub




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

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    

        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(LINENO)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
            
                If Trim(waResult(inRow, LINENO)) = "" And _
                   Trim(waResult(inRow, DocNo)) = "" And _
                   Trim(waResult(inRow, CTLPRD)) = "" And _
                   Trim(waResult(inRow, NOOFITM)) = "" And _
                   Trim(waResult(inRow, ADJQTY)) = "" And _
                   Trim(waResult(inRow, ADJAMT)) = "" And _
                   Trim(waResult(inRow, UPDATE)) = "" And _
                   Trim(waResult(inRow, DocID)) = "" Then
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
        
        If waResult(LastRow, UPDATE) = "N" Then
            .Col = UPDATE
            .Row = LastRow
            gsMsg = "Must be updated!"
            MsgBox gsMsg, vbOKOnly, gsTitle
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
    Dim adcmdSave As New ADODB.Command
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
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SCT001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, Set_MedDate(medDocDate))
    Call SetSPPara(adcmdSave, 6, "")
    Call SetSPPara(adcmdSave, 7, 0)
    Call SetSPPara(adcmdSave, 8, wlSaleID)
    
    Call SetSPPara(adcmdSave, 9, wsSrcCode)
    Call SetSPPara(adcmdSave, 10, txtRmk.Text)
    Call SetSPPara(adcmdSave, 11, wsFormID)
    
    Call SetSPPara(adcmdSave, 12, gsUserID)
    Call SetSPPara(adcmdSave, 13, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 14)
    wsDocNo = GetSPPara(adcmdSave, 15)
    
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

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.medDocDate.Enabled = False
            
            Me.cboSaleCode.Enabled = False
            Me.cboWhsCodeFr.Enabled = False
            Me.txtRmk.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            Me.medDocDate.Enabled = True
            Me.cboSaleCode.Enabled = True
    
            Me.cboWhsCodeFr.Enabled = True
            Me.txtRmk.Enabled = True
            
            
       '     If wiAction <> AddRec Then
                Me.tblDetail.Enabled = True
       '     End If
    End Select
End Sub

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
    
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "SCHDDOCNO"
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
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "文件號"
    vFilterAry(1, 2) = "SCHDDOCNO"
    
    vFilterAry(2, 1) = "日期"
    vFilterAry(2, 2) = "SCHDDOCDATE"
    
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "文件號"
    vAry(1, 2) = "SCHDDOCNO"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "日期"
    vAry(2, 2) = "SCHDDOCDATE"
    vAry(2, 3) = "1500"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT SCHDDOCNO, SCHDDOCDATE "
        wsSQL = wsSQL + "FROM ICSTKCNT "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE SCHDStatus = '1' "
        .sBindOrderSQL = "ORDER BY SCHDDocNo"
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
        
        If Chk_cboSaleCode() = True Then
            cboWhsCodeFr.SetFocus
        End If
    End If
End Sub

Private Sub cboSaleCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode

        
    wsSQL = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "AND SaleType = 'W' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    
    Call Ini_Combo(2, wsSQL, cboSaleCode.Left, cboSaleCode.Top + cboSaleCode.Height, tblCommon, wsFormID, "TBLSALECOD", Me.Width, Me.Height)
    
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

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Dim wsCurRecLn2 As String
    Dim wsCurRecLn3 As String
    
    Chk_NoDup = False
    
    wsCurRecLn = tblDetail.Columns(DocNo)
    'wsCurRecLn2 = tblDetail.Columns(WHSCODE)
    'wsCurRecLn3 = tblDetail.Columns(LOTNO)
   
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           'If wsCurRecLn = waResult(wlCtr, ITMCODE) And _
              wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
              wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
            If wsCurRecLn = waResult(wlCtr, DocNo) Then
              gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
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
              wsCurRecLn = waResult(wlCtr, ITMCODE) And _
              wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
              wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
            If wsCurRec = waResult(wlCtr, DocNo) Then
                gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
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
    
    
End Sub

Private Sub mnuPopUpSub_Click(Index As Integer)
    Call Call_PopUpMenu(waPopUpSub, Index)
End Sub

Private Sub Call_PopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

    Dim wsAct As String
    
    wsAct = inArray(inMnuIdx, 0)
    
    With tblDetail
    
    End With
End Sub

Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcPost).Enabled = False
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
                .Buttons(tcPost).Enabled = False
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
                .Buttons(tcPost).Enabled = True
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
                .Buttons(tcPost).Enabled = False
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
    Dim wsSQL As String
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
    
    
    wsSQL = "EXEC usp_RPTSTKCNT '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'盤點數目報表', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(30, "z") & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & gsLangID
        
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTSTKCNT"
    Else
    wsRptName = "RPTSTKCNT"
    End If
    
    
    NewfrmPrint.ReportID = "STKCNT"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "STKCNT"
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

    
    If waResult.UpperBound(1) >= 0 Then
        
   
        tblDetail.ReBind
        tblDetail.FirstRow = 0
        
        Call Calc_Total
    End If
End Sub


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property

Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property

Public Property Let SrcCode(InSrcCode As String)
    wsSrcCode = InSrcCode
End Property
Public Property Let TrnType(InTrnType As String)
    wsTrnType = InTrnType
End Property

Private Function Calc_Total(Optional ByVal LastRow As Variant) As Boolean
    Dim wdTotalQty As Double
    Dim wdTotalPrice As Double
    
    Dim wiRowCtr As Integer
    
    Calc_Total = False
    
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wdTotalQty = wdTotalQty + To_Value(waResult(wiRowCtr, ADJQTY))
        wdTotalPrice = wdTotalPrice + To_Value(waResult(wiRowCtr, ADJAMT))
    Next
    
    lblDspTrnAmt.Caption = Format(CStr(wdTotalPrice), gsQtyFmt)
    lblDspTrnQty.Caption = Format(CStr(wdTotalQty), gsQtyFmt)
    
    Calc_Total = True

End Function

Private Function Chk_cboWhsCodeFr() As Boolean
    Dim wsStatus As String

    Chk_cboWhsCodeFr = False
    
    If Trim(cboWhsCodeFr.Text) = "" Then
        gsMsg = "沒有輸入來源倉庫!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWhsCodeFr.SetFocus
        Exit Function
    End If

    If Chk_WhsCode(cboWhsCodeFr.Text, wsStatus) = False Then
        gsMsg = "倉庫不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboWhsCodeFr.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "倉庫已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboWhsCodeFr.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboWhsCodeFr = True
End Function

Private Function Chk_WhsCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_WhsCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT WhsStatus "
    wsSQL = wsSQL & " FROM MstWareHouse WHERE WhsCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "WhsStatus")
    
    Chk_WhsCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function




Private Sub txtRmk_GotFocus()
    FocusMe txtRmk
End Sub

Private Sub txtRmk_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtRmk, 50, KeyAscii)

    If KeyAscii = vbKeyReturn Then
    
            tblDetail.SetFocus
        
    End If
End Sub

Private Sub txtRmk_LostFocus()
    FocusMe txtRmk, True
End Sub



Private Function cmdPost() As Boolean
    
    Dim wsGenDte As String
    Dim adcmdPost As New ADODB.Command
    Dim wiRetKey As Integer
    Dim wsDteTim As String
    
    
    
    On Error GoTo cmdPost_Err
    
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
   
    If PostValidation() = False Then
       MousePointer = vbDefault
       Exit Function
    End If
    
    
    cnCon.BeginTrans
    Set adcmdPost.ActiveConnection = cnCon
        
    adcmdPost.CommandText = "USP_SOP000A_SK"
    adcmdPost.CommandType = adCmdStoredProc
    adcmdPost.Parameters.Refresh
    
    Call SetSPPara(adcmdPost, 1, wiAction)
    Call SetSPPara(adcmdPost, 2, wsTrnCd)
    Call SetSPPara(adcmdPost, 3, wlKey)
    Call SetSPPara(adcmdPost, 4, "")
    Call SetSPPara(adcmdPost, 5, gsUserID)
    Call SetSPPara(adcmdPost, 6, wsGenDte)
    Call SetSPPara(adcmdPost, 7, wsDteTim)
    Call SetSPPara(adcmdPost, 8, gsLangID)
    
    adcmdPost.Execute
                
    wiRetKey = GetSPPara(adcmdPost, 9)
                
    If wiRetKey <> 1 Then
         gsMsg = "Post Falied!"
         MsgBox gsMsg, vbOKOnly, gsTitle
         GoTo cmdPost_Err
    Else
         gsMsg = "Post Success!"
         MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    
    cnCon.CommitTrans
    
    Call cmdCancel
    Set adcmdPost = Nothing
    cmdPost = True
    
    MousePointer = vbDefault
    
    Exit Function
    
cmdPost_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdPost = Nothing
    
End Function


Private Function PostValidation() As Boolean
    
    PostValidation = False
    
    On Error GoTo PostValidation_Err
    
    If Not Chk_medDocDate Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    
    If Not Chk_cboWhsCodeFr Then Exit Function
    
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, DocNo)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = DocNo
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
            tblDetail.Col = DocNo
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    
    PostValidation = True
    
    Exit Function
    
PostValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
