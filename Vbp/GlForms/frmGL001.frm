VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGL001 
   Caption         =   "訂貨單"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "frmGL001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "frmGL001.frx":030A
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtRemark 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Text            =   "01234567890123457890"
      Top             =   1200
      Width           =   5865
   End
   Begin VB.ComboBox cboPfx 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtRevNo 
      Height          =   324
      Left            =   5880
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "12345678901234567890"
      Top             =   480
      Width           =   408
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin MSMask.MaskEdBox medDocDate 
      Height          =   285
      Left            =   1320
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
      Left            =   7800
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
            Picture         =   "frmGL001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL001.frx":66AD
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
         NumButtons      =   13
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
            Key             =   "Print"
            ImageIndex      =   13
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "frmGL001.frx":69D1
      TabIndex        =   5
      Top             =   1560
      Width           =   11535
   End
   Begin VB.Label lblRemark 
      Caption         =   "REMARK"
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label lblBalAmtLoc 
      Caption         =   "NETAMTLOC"
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   8100
      Width           =   1755
   End
   Begin VB.Label lblDspBalAmtLoc 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "9.999.999.999.99"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10320
      TabIndex        =   13
      Top             =   8100
      Width           =   1290
   End
   Begin VB.Label lblCtlPrd 
      Caption         =   "CTLPRD"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lblDspCtlPrd 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   5880
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblDocDate 
      Caption         =   "DOCDATE"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label lblRevNo 
      Caption         =   "REVNO"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   540
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
      TabIndex        =   6
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
Attribute VB_Name = "frmGL001"
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




Private wsOldCusNo As String
Private wsOldCurCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String




Private Const GACCCODE = 0
Private Const GACCNAME = 1
Private Const GJOBNO = 2
Private Const GCURR = 3
Private Const GEXCR = 4
Private Const GDAMT = 5
Private Const GCAMT = 6
Private Const GTAMT = 7
Private Const GCHQNO = 8
Private Const GCHQDATE = 9
Private Const GRMK = 10
Private Const GACCID = 11


Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"
Private Const tcPrint = "Print"

Private wiOpenDoc As Integer
Private wiAction As Integer
Private wiRevNo As Integer


Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "GLVOHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsSrcCd As String

Private wsDocNo As String


Private wbErr As Boolean
Private wsBaseCurCd As String
Private wsBaseExcr As String
Private wsCurrFlg As Boolean
Private wsSOPFlg As String
Private wsTitle As String
    
Private wbLock As Boolean
Private wbReadOnly As Boolean



Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, GACCCODE, GACCID
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
    
    wlKey = 0
  
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False

    
    Me.Caption = wsFormCaption
    
    wbLock = False
    wbReadOnly = False
    Call Ini_UnLockGrid
    
    FocusMe cboPfx
 
    

End Sub

Private Sub txtRemark_GotFocus()
    FocusMe txtRemark
End Sub

Private Sub txtRemark_LostFocus()
FocusMe txtRemark, True
End Sub



Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Dim wsExcRate As String
    Dim wsExcDesc As String
    
    Call chk_InpLen(txtRemark, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
           If Chk_KeyFld Then
            tblDetail.SetFocus
           End If
           
        
    End If
    
End Sub




Private Sub cboDocNo_GotFocus()
    
    FocusMe cboDocNo

End Sub

Private Sub cboDocNo_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
        
    wsSQL = "SELECT VOHDDOCNO, VOHDDOCDATE "
    wsSQL = wsSQL & " FROM GLVOHD "
    wsSQL = wsSQL & " WHERE VOHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND VOHDPFX LIKE '%" & IIf(cboPfx.SelLength > 0, "", Set_Quote(cboPfx.Text)) & "%' "
    wsSQL = wsSQL & " AND VOHDSTATUS = '1'"
    wsSQL = wsSQL & " ORDER BY VOHDDOCNO DESC "
  
    
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
        
        If Chk_cboPfx() = False Then Exit Sub
        
        If Chk_cboDocNo() = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If

End Sub

Private Function Chk_cboDocNo() As Boolean
Dim wsStatus As String
Dim wsPgmNo As String
Dim wsDocDate As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoVou(cboPfx) = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNo.SetFocus
        Exit Function
    End If
    
        
   If Chk_DocNo(cboDocNo, cboPfx, wsStatus, wsPgmNo, wsDocDate) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
        End If
        
        If Chk_ValidDocDate(wsDocDate, "GL") = False Then
            wbReadOnly = True
        End If

      '  If wsPgmNo <> wsFormID Then
      '  Call Ini_LockGrid
      '  wbLock = True
      '  End If
    
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
       
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
     Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    
    medDocDate.SetFocus
        
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
       
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
         Case vbKeyF9
        
            If tbrProcess.Buttons(tcPrint).Enabled = True Then Call cmdPrint
            
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
    
   
    wsSQL = "SELECT VOHDDOCID, VOHDDOCNO, VOHDDOCDATE, VOHDREVNO, "
    wsSQL = wsSQL & "VOHDREMARK, VODTACCID, VOHDCTLPRD, VODTREMARK, VODTCHQDATE, VODTCHQNO,"
    wsSQL = wsSQL & "COAACCCODE, VODTDESC, VODTJOBNO, VODTCURR, VODTEXCR, VODTDRAMT, VODTCRAMT, VODTTRNAMTL "
    wsSQL = wsSQL & "FROM  GLVOHD, GLVODT, mstCOA "
    wsSQL = wsSQL & "WHERE VOHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
    wsSQL = wsSQL & "AND VOHDPFX = '" & Set_Quote(cboPfx) & "' "
    wsSQL = wsSQL & "AND VOHDDOCID = VODTDOCID "
    wsSQL = wsSQL & "AND VODTACCID = COAACCID "
    wsSQL = wsSQL & "AND VOHDSTATUS <> '2'"
    wsSQL = wsSQL & "ORDER BY VODTDOCLINE "
  
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsRcd, "VOHDDOCID")
    txtRevNo.Text = Format(ReadRs(rsRcd, "VOHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsRcd, "VOHDREVNO"))
    medDocDate.Text = ReadRs(rsRcd, "VOHDDOCDATE")
    txtRemark.Text = ReadRs(rsRcd, "VOHDREMARK")
    lblDspCtlPrd.Caption = Right(ReadRs(rsRcd, "VOHDCTLPRD"), 2) & "/" & Left(ReadRs(rsRcd, "VOHDCTLPRD"), 4)
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, GACCCODE, GACCID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), GACCCODE) = ReadRs(rsRcd, "COAACCCODE")
             waResult(.UpperBound(1), GACCNAME) = ReadRs(rsRcd, "VODTDESC")
             waResult(.UpperBound(1), GJOBNO) = ReadRs(rsRcd, "VODTJOBNO")
             waResult(.UpperBound(1), GCURR) = ReadRs(rsRcd, "VODTCURR")
             waResult(.UpperBound(1), GEXCR) = Format(ReadRs(rsRcd, "VODTEXCR"), gsExrFmt)
             waResult(.UpperBound(1), GDAMT) = Format(ReadRs(rsRcd, "VODTDRAMT"), gsAmtFmt)
             waResult(.UpperBound(1), GCAMT) = Format(ReadRs(rsRcd, "VODTCRAMT"), gsAmtFmt)
             waResult(.UpperBound(1), GTAMT) = Format(ReadRs(rsRcd, "VODTTRNAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), GACCID) = To_Value(ReadRs(rsRcd, "VODTACCID"))
             waResult(.UpperBound(1), GRMK) = ReadRs(rsRcd, "VODTREMARK")
             waResult(.UpperBound(1), GCHQNO) = ReadRs(rsRcd, "VODTCHQNO")
             waResult(.UpperBound(1), GCHQDATE) = ReadRs(rsRcd, "VODTCHQDATE")
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
    lblRemark.Caption = Get_Caption(waScrItm, "REMARK")
    lblCtlPrd.Caption = Get_Caption(waScrItm, "CTLPRD")
    lblBalAmtLoc.Caption = Get_Caption(waScrItm, "BALAMTLOC")
    wsTitle = Get_Caption(waScrItm, "RPTTITLE")
    
    With tblDetail
        .Columns(GACCCODE).Caption = Get_Caption(waScrItm, "GACCCODE")
        .Columns(GACCNAME).Caption = Get_Caption(waScrItm, "GACCNAME")
        .Columns(GJOBNO).Caption = Get_Caption(waScrItm, "GJOBNO")
        .Columns(GCURR).Caption = Get_Caption(waScrItm, "GCURR")
        .Columns(GEXCR).Caption = Get_Caption(waScrItm, "GEXCR")
        .Columns(GDAMT).Caption = Get_Caption(waScrItm, "GDAMT")
        .Columns(GCAMT).Caption = Get_Caption(waScrItm, "GCAMT")
        .Columns(GTAMT).Caption = Get_Caption(waScrItm, "GTAMT")
        .Columns(GRMK).Caption = Get_Caption(waScrItm, "GRMK")
        .Columns(GCHQNO).Caption = Get_Caption(waScrItm, "GCHQNO")
        .Columns(GCHQDATE).Caption = Get_Caption(waScrItm, "GCHQDATE")
    End With
    
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "GLADD")
    wsActNam(2) = Get_Caption(waScrItm, "GLEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "GLDELETE")
    
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
 '   If Me.WindowState = 0 Then
 '       Me.Height = 9000
 '       Me.Width = 12000
 '   End If
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
    Set frmGL001 = Nothing

End Sub







Private Sub medDocDate_GotFocus()
    
  FocusMe medDocDate
    
End Sub


Private Sub medDocDate_LostFocus()

    FocusMe medDocDate, True
    
End Sub


Private Sub medDocDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medDocDate Then txtRemark.SetFocus
    End If
End Sub

Private Function Chk_medDocDate() As Boolean
Dim wsCtrlPrd As String
    
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
    
    
    If Chk_ValidDocDate(medDocDate.Text, "GL") = False Then
        medDocDate.SetFocus
        Exit Function
    End If
    
     wsCtrlPrd = Get_FiscalPeriod(medDocDate.Text)
    lblDspCtlPrd.Caption = Right(wsCtrlPrd, 2) & "/" & Left(wsCtrlPrd, 4)

    
    Chk_medDocDate = True

End Function




Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case GACCCODE
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
              Case GACCCODE
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

    
    wsSQL = "SELECT VOHDSTATUS FROM GLVOHD WHERE VOHDDOCNO = '" & Set_Quote(cboDocNo) & "' AND VOHDPFX = '" & Set_Quote(cboPfx) & "'"
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
    
    
    If Chk_medDocDate = False Then
        Exit Function
    End If
    
    tblDetail.Enabled = True
    Chk_KeyFld = True

End Function
Private Function cmdSave() As Boolean
    Dim wsDteTim As String
    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlRowCtr As Long
    Dim wsCtlPrd As String
    Dim wsSts As String
    Dim i As Integer
    Dim wdTmpAmt As Double
     
    On Error GoTo cmdSave_Err
    
    wsDteTim = Change_SQLDate(Now)
    
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
        
    adcmdSave.CommandText = "USP_GL001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, Trim(cboPfx.Text))
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, medDocDate.Text)
    Call SetSPPara(adcmdSave, 6, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 7, txtRemark.Text)
    Call SetSPPara(adcmdSave, 8, wsFormID)
    Call SetSPPara(adcmdSave, 9, gsUserID)
    Call SetSPPara(adcmdSave, 10, wsGenDte)
    Call SetSPPara(adcmdSave, 11, wsDteTim)
    
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 12)
    wsDocNo = GetSPPara(adcmdSave, 13)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_GL001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GACCCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, GACCID))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, GACCNAME))
                Call SetSPPara(adcmdSave, 5, wiCtr + 1)
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, GJOBNO))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, GCURR))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, GEXCR))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, GDAMT))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, GCAMT))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, GTAMT))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, GRMK))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, GCHQNO))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, GCHQDATE))
                Call SetSPPara(adcmdSave, 15, wsFormID)
                Call SetSPPara(adcmdSave, 16, gsUserID)
                Call SetSPPara(adcmdSave, 17, wsGenDte)
                
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
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, GACCCODE)) <> "" Then
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
    
    
    If To_Value(lblDspBalAmtLoc.Caption) <> 0 Then
        gsMsg = "Balance Amount must equal to zero!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
    


Private Sub cmdNew()

    Dim newForm As New frmGL001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmGL001
    
    newForm.OpenDoc = True
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show

End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
  '  Me.Left = (Screen.Width - Me.Width) / 2
  '  Me.Top = (Screen.Height - Me.Height) / 2
     
    Me.WindowState = 2
     
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "GL001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    Call getExcRate(wsBaseCurCd, gsSystemDate, wsBaseExcr, "")
    
    wsSrcCd = "GL"
    wsTrnCd = "62"
    
    If wsBaseCurCd <> "" Then
    wsCurrFlg = False
    Else
    wsCurrFlg = True
   End If
    
    wsSOPFlg = Get_SystemFlag("SYPINTSOP")

End Sub



Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboPfx.SetFocus
    
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
 Dim wsPrtDocNo As String
 Dim wsPrtPfx As String
 
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
        Case tcPrint
           If MsgBox("你是否確定儲存現時之變更而列印?", vbYesNo, gsTitle) = vbYes Then
                wsPrtDocNo = cboDocNo.Text
                wsPrtPfx = cboPfx.Text
                If cmdSave = False Then Exit Sub
                cboDocNo.Text = wsPrtDocNo
                cboPfx.Text = wsPrtPfx
                Call Ini_Scr_AfrKey
           End If
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


Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = GACCCODE To GACCID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case GACCCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                Case GACCNAME
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 50
                Case GJOBNO
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                Case GCURR
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Visible = wsCurrFlg
                Case GEXCR
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsExrFmt
                    .Columns(wiCtr).Visible = wsCurrFlg
                Case GDAMT
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GCAMT
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GTAMT
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case GRMK
                    .Columns(wiCtr).Width = 6000
                    .Columns(wiCtr).DataWidth = 50
                Case GCHQNO
                    .Columns(wiCtr).Width = 4000
                    .Columns(wiCtr).DataWidth = 15
                Case GCHQDATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                       
                Case GACCID
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
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
    Dim wlAccID As Long
    Dim wsDes As String
    Dim wsExcRat As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case GACCCODE
            
                If Chk_Lock Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdAccCode(.Columns(ColIndex).Text, wlAccID, wsDes) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(GACCID).Text = wlAccID
                .Columns(GACCNAME).Text = wsDes
                .Columns(GCURR).Text = wsBaseCurCd
                .Columns(GEXCR).Text = To_Value(wsBaseExcr)
                
                
             Case GCURR
                
                If Chk_grdCurr(.Columns(ColIndex).Text, wsExcRat) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(GEXCR).Text = NBRnd(To_Value(wsExcRat), giExrDp)
            
            Case GEXCR
                
                If chk_grdExcRat(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case GJOBNO
                
                If Chk_grdJobNo(.Columns(ColIndex).Text) = False Then
                     GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case GCAMT, GDAMT
            
                If Chk_Lock Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                                
                If Chk_Amount(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                If ColIndex = GDAMT Then
                    .Columns(GDAMT).Text = Format(.Columns(GDAMT).Text, "#,##0." & String(giAmtDp, "0"))
                    .Columns(GCAMT).Text = ""
                Else
                    .Columns(GCAMT).Text = Format(.Columns(GCAMT).Text, "#,##0." & String(giAmtDp, "0"))
                    .Columns(GDAMT).Text = ""
                End If
                
            End Select

'            If To_Value(.Columns(GDAMT).Text) > 0 Then
'            .Columns(GTAMT).Text = NBRnd(To_Value(.Columns(GDAMT).Text) * _
'                                             To_Value(.Columns(GEXCR).Text), giAmtDp)
'            ElseIf To_Value(.Columns(GCAMT).Text) > 0 Then
'            .Columns(GTAMT).Text = NBRnd(0 - (To_Value(.Columns(GCAMT).Text) * _
'                                             To_Value(.Columns(GEXCR).Text)), giAmtDp)
'            Else
'            .Columns(GTAMT).Text = Format(To_Value(.Columns(GTAMT).Text), "#,##0." & String(giAmtDp, "0"))
'            End If

            .Columns(GTAMT).Text = NBRnd((To_Value(.Columns(GDAMT).Text) - To_Value(.Columns(GCAMT).Text)) * _
                                             To_Value(.Columns(GEXCR).Text), giAmtDp)
            
            
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
    Dim wsCtlDte As String
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case GACCCODE
                
                wsSQL = "SELECT COAACCCODE, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " FROM mstCOA "
                wsSQL = wsSQL & " WHERE COASTATUS <> '2' "
                wsSQL = wsSQL & " AND COAACCCODE LIKE '%" & IIf(Trim(.SelText) <> "", "", Set_Quote(.Columns(GACCCODE).Text)) & "%' "
                wsSQL = wsSQL & " AND COAACCID NOT IN ( "
                wsSQL = wsSQL & " SELECT COAACCID ACCID FROM MSTCOMPANY, MSTCOA WHERE CMPID = '01' AND COAACCCODE = CMPCURREARN "
                wsSQL = wsSQL & " ) "
                
                wsSQL = wsSQL & " ORDER BY COAACCCODE "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLCOA", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case GCURR
                
            wsCtlDte = IIf(Trim(medDocDate.Text) = "" Or Trim(medDocDate.Text) = "/  /", gsSystemDate, medDocDate.Text)
            wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(Trim(.SelText) <> "", "", Set_Quote(.Columns(GCURR).Text)) & "%' "
            wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
            wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
            wsSQL = wsSQL & " AND EXCSTATUS = '1' "
            wsSQL = wsSQL & "ORDER BY EXCCURR "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case GJOBNO
            
            If wsSOPFlg = "Y" Then
                
                wsSQL = "SELECT SOHDDOCNO, CUSCODE FROM SOASOHD, MSTCUSTOMER "
                wsSQL = wsSQL & " WHERE SOHDSTATUS = '4' "
                wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(GJOBNO).Text) & "%' "
                wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
                wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
            Else
                wsSQL = "SELECT JOBCODE, JOBNAME FROM mstJOB "
                wsSQL = wsSQL & " WHERE JOBSTATUS <> '2' AND JOBCODE LIKE '%" & Set_Quote(.Columns(GJOBNO).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY JOBCODE "
                
            End If
            
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLJOBCODE", Me.Width, Me.Height)
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
            If Chk_Lock Then Exit Sub
            If .Bookmark = waResult.UpperBound(1) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case vbKeyF8        ' DELETE LINE
            KeyCode = vbDefault
            If Chk_Lock Then Exit Sub
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
                Case GACCCODE, GACCNAME, GCURR, GEXCR, GDAMT, GCAMT, GTAMT, GCHQNO, GCHQDATE
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case GJOBNO
                    KeyCode = vbDefault
                    
                    If wsCurrFlg = False Then
                    .Col = GDAMT
                    Else
                    .Col = GCURR
                    End If
                    
                Case GRMK
                    KeyCode = vbKeyDown
                    .Col = GACCCODE
            End Select
        Case vbKeyLeft
            Select Case .Col
                Case GACCNAME, GJOBNO, GCURR, GEXCR, GCAMT, GTAMT, GCHQNO, GCHQDATE, GRMK
                    KeyCode = vbDefault
                    .Col = .Col - 1
                Case GDAMT
                    KeyCode = vbDefault
                    
                    If wsCurrFlg = False Then
                    .Col = GJOBNO
                    Else
                    .Col = GEXCR
                    End If
                    
            End Select
            
        Case vbKeyRight
            Select Case .Col
                Case GACCCODE, GACCNAME, GCURR, GEXCR, GDAMT, GCAMT, GTAMT, GCHQNO, GCHQDATE
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case GJOBNO
                    KeyCode = vbDefault
                    
                    If wsCurrFlg = False Then
                    .Col = GDAMT
                    Else
                    .Col = GCURR
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
        
        Case GDAMT, GCAMT
            Call Chk_InpNum(KeyAscii, tblDetail.Text, True, True)
        
      
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = GACCCODE
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case GACCCODE
                    Call Chk_grdAccCode(.Columns(GACCCODE), 0, "")
                Case GCURR
                    Call Chk_grdCurr(.Columns(GCURR).Text, "")
                Case GEXCR
                    Call chk_grdExcRat(.Columns(GEXCR).Text)
                 Case GJOBNO
                    Call Chk_grdJobNo(.Columns(GJOBNO).Text)
                Case GCAMT, GDAMT
                    Call Chk_Amount(.Columns(.Col).Text)
                
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub



Private Function Chk_grdAccCode(inNo As String, ByRef outAccID As Long, ByRef OutDesc As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdAccCode = False
    
    If Trim(inNo) = "" Then
        Chk_grdAccCode = True
        Exit Function
    End If
    
    wsSQL = "SELECT COAACCID, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " DES FROM mstCOA"
    wsSQL = wsSQL & " WHERE COAAccCode = '" & Set_Quote(inNo) & "' "
    wsSQL = wsSQL & " AND COASTATUS = '1' "
    wsSQL = wsSQL & " AND COAACCID NOT IN ( "
    wsSQL = wsSQL & " SELECT COAACCID ACCID FROM MSTCOMPANY, MSTCOA WHERE CMPID = '01' AND COAACCCODE = CMPCURREARN "
    wsSQL = wsSQL & " ) "
       
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        gsMsg = "No Such Account Code!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outAccID = ReadRs(rsRcd, "COAACCID")
    OutDesc = ReadRs(rsRcd, "DES")
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_grdAccCode = True
        

End Function



Private Function Chk_grdCurr(inNo As String, ByRef outExcr As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdCurr = False
    
    If Trim(inNo) = "" Then
        gsMsg = "Curreny Must Input!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If getExcRate(inNo, medDocDate.Text, outExcr, "") = False Then
            gsMsg = "No Such Curreny in This Month!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
    End If
        
    
    Chk_grdCurr = True
        

End Function


Private Function chk_grdExcRat(inRate As String) As Boolean
    
    chk_grdExcRat = False
    
    If To_Value(inRate) = 0 Then
        gsMsg = "Can not equal to Zero!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
    If To_Value(inRate) > To_Value(gsMaxVal) Then
       gsMsg = "數量太大!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
    chk_grdExcRat = True
    
End Function


Private Function Chk_Amount(inAmt As String) As Integer
    
    Chk_Amount = False
    
    If Trim(inAmt) = "" Then
        Chk_Amount = True
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
                If Trim(.Columns(GACCCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, GACCCODE)) = "" And _
                   Trim(waResult(inRow, GACCNAME)) = "" And _
                   Trim(waResult(inRow, GJOBNO)) = "" And _
                   Trim(waResult(inRow, GCURR)) = "" And _
                   Trim(waResult(inRow, GEXCR)) = "" And _
                   Trim(waResult(inRow, GDAMT)) = "" And _
                   Trim(waResult(inRow, GCAMT)) = "" And _
                   Trim(waResult(inRow, GTAMT)) = "" And _
                   Trim(waResult(inRow, GRMK)) = "" And _
                   Trim(waResult(inRow, GCHQNO)) = "" And _
                   Trim(waResult(inRow, GCHQDATE)) = "" And _
                   Trim(waResult(inRow, GACCID)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
End Function


Private Function Chk_grdJobNo(inNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdJobNo = False
    
    'If Trim(inNo) = "" Then
        Chk_grdJobNo = True
        Exit Function
    'End If
    
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
        
        If Chk_grdAccCode(waResult(LastRow, GACCCODE), 0, "") = False Then
            .Col = GACCCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdCurr(waResult(LastRow, GCURR), "") = False Then
                .Col = GCURR
                .Row = LastRow
                Exit Function
        End If
        
        
        
        If chk_grdExcRat(waResult(LastRow, GEXCR)) = False Then
                .Col = GEXCR
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdJobNo(waResult(LastRow, GJOBNO)) = False Then
                .Col = GJOBNO
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_Amount(waResult(LastRow, GDAMT)) = False Then
            .Col = GDAMT
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_Amount(waResult(LastRow, GCAMT)) = False Then
            .Col = GCAMT
            .Row = LastRow
            Exit Function
        End If
        
         If To_Value(waResult(LastRow, GCAMT)) = 0 And To_Value(waResult(LastRow, GDAMT)) = 0 Then
            gsMsg = "Must Have Amount in CR or DR side!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            .Col = GDAMT
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
        wiTotal = wiTotal + To_Value(waResult(wiRowCtr, GTAMT))
    Next
    
    lblDspBalAmtLoc.Caption = Format(CStr(wiTotal), gsAmtFmt)
    
    Calc_Total = True

End Function




Private Function cmdDel() As Boolean
    Dim wsDteTim As String
    Dim wsGenDte As String
    Dim adcmdDelete As New ADODB.Command
    Dim wsDocNo As String
    Dim i As Integer
    
    cmdDel = False
    
    MousePointer = vbHourglass
    
    On Error GoTo cmdDelete_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    wsDteTim = Change_SQLDate(Now)
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbLock Or wbReadOnly Then
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
        
    adcmdDelete.CommandText = "USP_GL001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wlKey)
    Call SetSPPara(adcmdDelete, 3, Trim(cboPfx.Text))
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 6, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 7, txtRemark.Text)
    Call SetSPPara(adcmdDelete, 8, wsFormID)
    Call SetSPPara(adcmdDelete, 9, gsUserID)
    Call SetSPPara(adcmdDelete, 10, wsGenDte)
    Call SetSPPara(adcmdDelete, 11, wsDteTim)
    
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
                .Buttons(tcPrint).Enabled = False
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
                .Buttons(tcPrint).Enabled = False
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
                .Buttons(tcPrint).Enabled = False
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
                .Buttons(tcPrint).Enabled = False
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
                .Buttons(tcPrint).Enabled = True
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "ReadOnly"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcPrint).Enabled = True
                .Buttons(tcExit).Enabled = True
            
            End With
            
       
    
    End Select
End Sub



'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
        
            Me.cboPfx.Enabled = False
            Me.cboDocNo.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            Me.txtRemark.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboPfx.Enabled = True
            Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboPfx.Enabled = True
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboPfx.Enabled = False
            Me.cboDocNo.Enabled = False
            
            Me.txtRevNo.Enabled = True
            Me.medDocDate.Enabled = True
            Me.txtRemark.Enabled = True
            
            
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
        .TableType = cboPfx.Text
        .TableKey = "VOHDDOCNO"
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
    vFilterAry(1, 1) = "Doc No."
    vFilterAry(1, 2) = "vohdDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "vohdDocDate"
    
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "vohdDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "voHdDocDate"
    vAry(2, 3) = "1500"
    
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT INHdDocNo, InHdDocDate "
        wsSQL = wsSQL + "FROM GLVOHD "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE InHdStatus = '1' "
        .sBindOrderSQL = "ORDER BY voHdDocNo"
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



Private Sub txtRevNo_LostFocus()
    FocusMe txtRevNo, True
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
           If Chk_Lock Then Exit Sub
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
            If Chk_Lock Then Exit Sub
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
Private Function Chk_DocNo(ByVal InDocNo As String, ByVal InPfx As String, ByRef OutStatus As String, ByRef OutPgmNo As String, ByRef OutDocDate As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    OutStatus = ""
    OutPgmNo = ""
    Chk_DocNo = False
    
    wsSQL = "SELECT VOHDDOCDATE, VOHDSTATUS, VOHDPGMNO FROM GLVOHD "
    wsSQL = wsSQL & " WHERE VOHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    wsSQL = wsSQL & " AND VOHDPFX = '" & Set_Quote(InPfx) & "'"
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount <= 0 Then
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    
    OutStatus = ReadRs(rsRcd, "VOHDSTATUS")
    OutPgmNo = ReadRs(rsRcd, "VOHDPGMNO")
    OutDocDate = ReadRs(rsRcd, "VOHDDOCDATE")
    
    rsRcd.Close
    Set rsRcd = Nothing
    
       
   
    Chk_DocNo = True
   

End Function



Private Sub cboPfx_GotFocus()
    
    FocusMe cboPfx

End Sub

Private Sub cboPfx_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPfx
  
    
    wsSQL = "SELECT VOUPREFIX, VOUDESC "
    wsSQL = wsSQL & " FROM SYSVOUNO "
    wsSQL = wsSQL & " WHERE VOUPREFIX LIKE '%" & IIf(cboPfx.SelLength > 0, "", Set_Quote(cboPfx.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY VOUPREFIX "
  
    
    Call Ini_Combo(2, wsSQL, cboPfx.Left, cboPfx.Top + cboPfx.Height, tblCommon, wsFormID, "TBLPFX", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub



Private Sub cboPfx_LostFocus()
FocusMe cboPfx, True
End Sub

Private Sub cboPfx_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLenC(cboPfx, 3, KeyAscii, True, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboPfx() = False Then Exit Sub
        
        cboDocNo.SetFocus
        
    End If

End Sub

Private Function Chk_cboPfx() As Boolean

Dim rsRcd As New ADODB.Recordset
Dim wsSQL As String
Dim wsStatus As String
Dim wsUpdFlg As String
Dim wsTrnCode As String
Dim wsDocDate As String
Dim wsPgmNo As String
    
    Chk_cboPfx = False
    
    If Trim(cboPfx.Text) = "" Then
        gsMsg = "Must Input Voucher Prefix!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboPfx.SetFocus
        Exit Function
    End If
    
        
 '  If Chk_VouPrefix(cboPfx.Text) = False Then
 '           gsMsg = "This is not a valid prefix!"
 '           MsgBox gsMsg, vbOKOnly, gsTitle
 '           cboPfx.SetFocus
 '           Exit Function
 '  End If
   Chk_cboPfx = True

End Function

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String

    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(5)
    wsSelection(1) = lblDocNo.Caption & " " & Set_Quote(cboDocNo.Text)

    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTGLP006 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wsTitle & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboPfx.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'000000', "
    wsSQL = wsSQL & "'999999', "
    wsSQL = wsSQL & "'', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & "'0000/00/00', "
    wsSQL = wsSQL & "'9999/99/99', "
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTGLP006P"
    Else
    wsRptName = "RPTGLP006P"
    End If
    
    NewfrmPrint.ReportID = "GLP006"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "GLP006"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
    
End Sub
Private Sub Ini_LockGrid()
    

    With tblDetail
        .EmptyRows = False
        .AllowAddNew = False
        .AllowDelete = False
        
        
    End With
    
End Sub

Private Sub Ini_UnLockGrid()
    

    With tblDetail
        .EmptyRows = True
        .AllowAddNew = True
        .AllowDelete = True
    End With
    
End Sub

Private Function Chk_Lock() As Boolean

    If wbLock Then
        gsMsg = "不能更改或刪除!文件由倉庫入數!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If

Chk_Lock = wbLock

End Function

                
