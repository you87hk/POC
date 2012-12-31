VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAR002 
   Caption         =   "訂貨單"
   ClientHeight    =   4155
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "frmAR002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "frmAR002.frx":030A
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtChqAmtOrg 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmAR002.frx":2A0D
      Top             =   3360
      Width           =   1900
   End
   Begin VB.ComboBox cboTMPML 
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   2010
   End
   Begin VB.TextBox txtRmk 
      Height          =   300
      Left            =   2040
      TabIndex        =   8
      Text            =   "012345678901234578901234567890123457890123456789"
      Top             =   3720
      Width           =   5535
   End
   Begin VB.ComboBox cboMLCode 
      Height          =   300
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   2010
   End
   Begin VB.ComboBox cboCurr 
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtExcr 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox cboCusCode 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin MSMask.MaskEdBox medDocDate 
      Height          =   285
      Left            =   8400
      TabIndex        =   1
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
      Left            =   6360
      Top             =   480
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
            Picture         =   "frmAR002.frx":2A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":32F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":3BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":4022
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":4474
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":478E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":4BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":5032
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":534C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":5666
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":5AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR002.frx":6394
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   14
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
   Begin VB.Label lblTmpMl 
      Caption         =   "MLCODE"
      Height          =   345
      Left            =   240
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblDspTmpMlDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4080
      TabIndex        =   26
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblChqAmtOrg 
      Caption         =   "Cheque Amount (Org)"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   3390
      Width           =   1695
   End
   Begin VB.Label lblRmk 
      Caption         =   "RMK"
      Height          =   240
      Left            =   240
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblDspMLDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4080
      TabIndex        =   23
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label lblMlCode 
      Caption         =   "MLCODE"
      Height          =   240
      Left            =   240
      TabIndex        =   22
      Top             =   2700
      Width           =   1695
   End
   Begin VB.Label lblCusTel 
      Caption         =   "CUSTEL"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label lblCusFax 
      Caption         =   "CUSFAX"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblDspCusFax 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   5520
      TabIndex        =   19
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblCusName 
      Caption         =   "CUSNAME"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label lblDspCusTel 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   2040
      TabIndex        =   17
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblExcr 
      Caption         =   "EXCR"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2340
      Width           =   1695
   End
   Begin VB.Label LblCurr 
      Caption         =   "CURR"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label lblDspCusName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lblCusCode 
      Caption         =   "CUSCODE"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblDocDate 
      Caption         =   "DOCDATE"
      Height          =   255
      Left            =   7125
      TabIndex        =   10
      Top             =   900
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
      Left            =   240
      TabIndex        =   9
      Top             =   540
      Width           =   1695
   End
End
Attribute VB_Name = "frmAR002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control




Private wsOldCusNo As String
Private wsOldCurCd As String
Private wsTmpMl As String
Private wsCtlDte As String
Private wbReadOnly As Boolean



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
Private wlCusID As Long


Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "ARCHEQUE"

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
    
    wsCtlDte = getCtrlMth("AR")
      
    
    wsOldCusNo = ""
    wsOldCurCd = ""
    wsTmpMl = ""
    
    wlKey = 0
    wlCusID = 0
    wbReadOnly = False
    
    tblCommon.Visible = False

    
    Me.Caption = wsFormCaption
    
    FocusMe cboDocNo

    

End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_LostFocus()
FocusMe cboCurr, True
End Sub

Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
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
           
            cboMLCode.SetFocus
           
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
  
    
    wsSQL = "SELECT ARCQCHQNO, ARCQCHQDATE, ARCQCURR, ARCQCHQAMT "
    wsSQL = wsSQL & " FROM ARCHEQUE "
    wsSQL = wsSQL & " WHERE ARCQCHQNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
  '  wsSql = wsSql & " AND ARCQUPDFLG = 'N'"
  '  wsSql = wsSql & " AND ARCQSTATUS = '1'"
    wsSQL = wsSQL & " AND ARCQPGMNO = '" & wsFormID & "' "
    wsSQL = wsSQL & " ORDER BY ARCQCHQNO, ARCQCHQDATE "
  
    
    Call Ini_Combo(4, wsSQL, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
Dim wsDocDate As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoGen("IV") = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNo.SetFocus
        Exit Function
    End If
    
        
   If Chk_DocNo(cboDocNo, wsStatus, wsUpdFlg, wsDocDate, wsPgmNo) = True Then
        
        
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
    
                
        If Chk_ValidDocDate(wsDocDate, "AR") = False Then
            cboDocNo.SetFocus
            Exit Function
        End If
        
        wsSQL = "SELECT ARCQCHQID FROM ARCHEQUE "
        wsSQL = wsSQL & " WHERE ARCQCHQNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND ARCQPGMNO <> '" & wsFormID & "' "
    
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
        rsRcd.Close
        Set rsRcd = Nothing
   
    Else
    
        wsSQL = "SELECT ARSHDOCID FROM ARSTHD "
        wsSQL = wsSQL & " WHERE ARSHDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND ARSHPGMNO <> '" & wsFormID & "' "
    
    
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
        
        wsSQL = "SELECT INHDDOCNO FROM ARINHD "
        wsSQL = wsSQL & " WHERE INHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & " AND INHDPGMNO <> '" & wsFormID & "' "
    
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount > 0 Then
            gsMsg = "Document No has been already used by Invocie!"
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
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
        wsOldCusNo = cboCusCode.Text
        wsOldCurCd = cboCurr.Text
        
    
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
    
    cboCusCode.SetFocus
        
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
    
   
    wsSQL = "SELECT ARCQCHQID, ARCQCHQNO, ARCQCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, "
    wsSQL = wsSQL & "ARCQCHQDATE, ARCQCURR, ARCQEXCR, ARCQBANKML, ARCQTMPML, "
    wsSQL = wsSQL & "ARCQCHQAMT, ARCQREMARK "
    wsSQL = wsSQL & "FROM  ARCHEQUE, mstCUSTOMER "
    wsSQL = wsSQL & "WHERE ARCQCHQNO = '" & Set_Quote(cboDocNo) & "' "
    wsSQL = wsSQL & "AND ARCQCUSID = CUSID "
    wsSQL = wsSQL & "AND ARCQPGMNO = '" & wsFormID & "'"
    wsSQL = wsSQL & "AND ARCQSTATUS <> '2'"
    
        
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    wlKey = ReadRs(rsRcd, "ARCQCHQID")
    medDocDate.Text = ReadRs(rsRcd, "ARCQCHQDATE")
    wlCusID = ReadRs(rsRcd, "CUSID")
    cboCusCode.Text = ReadRs(rsRcd, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsRcd, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsRcd, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsRcd, "CUSFAX")
    cboCurr.Text = ReadRs(rsRcd, "ARCQCURR")
    txtExcr.Text = Format(ReadRs(rsRcd, "ARCQEXCR"), gsExrFmt)
    txtChqAmtOrg.Text = NBRnd(ReadRs(rsRcd, "ARCQCHQAMT"), giAmtDp)
    cboMLCode = ReadRs(rsRcd, "ARCQBANKML")
    cboTMPML = ReadRs(rsRcd, "ARCQTMPML")
    txtRmk = ReadRs(rsRcd, "ARCQREMARK")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    rsRcd.Close
    
    
    wsSQL = " SELECT INHDMLCODE "
    wsSQL = wsSQL & " FROM  ARCHEQUE, ARINHD "
    wsSQL = wsSQL & " WHERE ARCQCHQID  = " & wlKey
    wsSQL = wsSQL & " AND INHDDOCNO  = ARCQCHQNO "
    wsSQL = wsSQL & " AND INHDPGMNO  = '" & wsFormID & "' "
        
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
       cboTMPML.Text = ReadRs(rsRcd, "INHDMLCODE")
    Else
       cboTMPML.Text = cboMLCode.Text
    End If
    
    lblDspTmpMlDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboTMPML.Text) & "'", "MLDESC")
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    lblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblTmpMl.Caption = Get_Caption(waScrItm, "TMPML")
    
    lblChqAmtOrg.Caption = Get_Caption(waScrItm, "CHQAMTORG")
    
    
    
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
        Me.Height = 4560
        Me.Width = 9915
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    Call UnLockAll(wsConnTime, wsFormID)
    Set waScrToolTip = Nothing
    Set waScrItm = Nothing
    Set frmAR002 = Nothing

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
    
    
    If Chk_ValidDocDate(medDocDate.Text, "AR") = False Then
        medDocDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medDocDate = True

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

    
    wsSQL = "SELECT ARCQSTATUS FROM ARCHEQUE WHERE ARCQCHQNO = '" & Set_Quote(cboDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

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
    
 
    
    
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_AR002"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 4, wlCusID)
    Call SetSPPara(adcmdSave, 5, medDocDate.Text)
    Call SetSPPara(adcmdSave, 6, cboCurr.Text)
    Call SetSPPara(adcmdSave, 7, txtExcr.Text)
    Call SetSPPara(adcmdSave, 8, To_Value(txtChqAmtOrg.Text))
    Call SetSPPara(adcmdSave, 9, NBRnd(To_Value(txtExcr.Text) * To_Value(txtChqAmtOrg.Text), giAmtDp))
    Call SetSPPara(adcmdSave, 10, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 11, cboTMPML.Text)
    Call SetSPPara(adcmdSave, 12, "")
    Call SetSPPara(adcmdSave, 13, txtRmk.Text)
    Call SetSPPara(adcmdSave, 14, wsFormID)
    Call SetSPPara(adcmdSave, 15, gsUserID)
    Call SetSPPara(adcmdSave, 16, wsGenDte)
    Call SetSPPara(adcmdSave, 17, wsDteTim)
    
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 18)
    wsDocNo = GetSPPara(adcmdSave, 19)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    
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
    If Not chk_cboCusCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboMLCode Then Exit Function
    If Not Chk_txtChqAmtOrg(txtChqAmtOrg.Text) Then Exit Function
    If Not Chk_cboTmpML Then Exit Function
    
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
    


Private Sub cmdNew()

    Dim newForm As New frmAR002
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmAR002
    
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
    wsFormID = "AR002"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    If getExcRate(wsBaseCurCd, wsConnTime, wsBaseExcr, "") = False Then
            wsBaseExcr = Format(1, gsExrFmt)
    End If
    wsSrcCd = "AR"

    
    


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
       
       txtRmk.SetFocus
       
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
        
                cboMLCode.SetFocus
       
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



Private Sub cboCusCode_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCode
    
    If gsLangID = "1" Then
        wsSQL = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSQL = wsSQL & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSQL = wsSQL & "AND CUSSTATUS = '1' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY CUSCODE "
    Else
        wsSQL = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSQL = wsSQL & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSQL = wsSQL & "AND CUSSTATUS = '1' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & "ORDER BY CUSCODE "
    End If
    Call Ini_Combo(2, wsSQL, cboCusCode.Left, cboCusCode.Top + cboCusCode.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
   
End Sub

Private Sub cboCusCode_GotFocus()
    
    Set wcCombo = cboCusCode
    'TREtoolsbar1.ButtonEnabled(tcCusSrh) = True
    FocusMe cboCusCode
    
End Sub

Private Sub cboCusCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_cboCusCode() = False Then Exit Sub
        If wiAction = AddRec Or wsOldCusNo <> cboCusCode.Text Then Call Get_DefVal
           
            cboCurr.SetFocus
            
    End If
    
End Sub

Private Function chk_cboCusCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    Dim wsTel As String
    Dim wsFax As String
    
    
    chk_cboCusCode = False
    
    
    If Trim(cboCusCode) = "" Then
        gsMsg = "必需輸入客戶編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
    
    If Chk_CusCode(cboCusCode, wlID, wsName, wsTel, wsFax) Then
        wlCusID = wlID
        lblDspCusName.Caption = wsName
        lblDspCusTel.Caption = wsTel
        lblDspCusFax.Caption = wsFax
    Else
        wlCusID = 0
        lblDspCusName.Caption = ""
        lblDspCusTel.Caption = ""
        lblDspCusFax.Caption = ""
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
    
    chk_cboCusCode = True

End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstCUSTOMER "
    wsSQL = wsSQL & "WHERE CUSID = " & wlCusID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "CUSCURR")
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

    
    gsMsg = "你是否確認要刪除此檔案?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       wiAction = CorRec
       MousePointer = vbDefault
       Exit Function
    End If
    
    wiAction = DelRec
    
    cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_AR002"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wlKey)
    Call SetSPPara(adcmdDelete, 3, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 4, wlCusID)
    Call SetSPPara(adcmdDelete, 5, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 6, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 7, txtExcr.Text)
    Call SetSPPara(adcmdDelete, 8, To_Value(txtChqAmtOrg.Text))
    Call SetSPPara(adcmdDelete, 9, NBRnd(To_Value(txtExcr.Text) * To_Value(txtChqAmtOrg.Text), giAmtDp))
    Call SetSPPara(adcmdDelete, 10, cboMLCode.Text)
    Call SetSPPara(adcmdDelete, 11, cboTMPML.Text)
    Call SetSPPara(adcmdDelete, 12, "")
    Call SetSPPara(adcmdDelete, 13, txtRmk.Text)
    Call SetSPPara(adcmdDelete, 14, wsFormID)
    Call SetSPPara(adcmdDelete, 15, gsUserID)
    Call SetSPPara(adcmdDelete, 16, wsGenDte)
    Call SetSPPara(adcmdDelete, 17, wsDteTim)
    
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 18)
    wsDocNo = GetSPPara(adcmdDelete, 19)
    
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
            Me.cboCusCode.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.cboTMPML.Enabled = False
            Me.txtRmk.Enabled = False
            Me.txtChqAmtOrg.Enabled = False
            
            
            
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            Me.cboCusCode.Enabled = True
            Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.cboMLCode.Enabled = True
            Me.cboTMPML.Enabled = True
            
            Me.txtRmk.Enabled = True
            Me.txtChqAmtOrg.Enabled = True
            
                        
            
    End Select
End Sub

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
    
        .TableID = wsKeyType
        .TableType = wsSrcCd
        .TableKey = "ARCQCHQNO"
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
    vFilterAry(1, 2) = "ARCQCHQNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "ARCQCHQDate"
    
    vFilterAry(3, 1) = "Customer #"
    vFilterAry(3, 2) = "CusCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "ARCQCHQNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "ARCQCHQDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Customer#"
    vAry(3, 2) = "CusCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Customer Name"
    vAry(4, 2) = "CusName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT ARCQCHQNo, ARCQCHQDate, mstCustomer.CusCode,  mstCustomer.CusName "
        wsSQL = wsSQL + "FROM MstCustomer, Archeque "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE ARCQStatus = '1' And ARCQCusID = CusID "
        .sBindOrderSQL = "ORDER BY ARCQDocNo"
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
        
        
        cboTMPML.SetFocus
       
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
        
        
        txtChqAmtOrg.SetFocus
       
    End If
    
End Sub

Private Sub cboTmpML_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboTMPML
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass "
    wsSQL = wsSQL & " WHERE MLCode LIKE '%" & IIf(cboTMPML.SelLength > 0, "", Set_Quote(cboTMPML.Text)) & "%' "
    wsSQL = wsSQL & " AND MLSTATUS = '1' "
    wsSQL = wsSQL & " AND MLTYPE = 'A' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboTMPML.Left, cboTMPML.Top + cboTMPML.Height, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboTmpML() As Boolean
Dim wsDesc As String

    Chk_cboTmpML = False
     
     
    If Trim(cboTMPML.Text) = "" Then
        gsMsg = "必需輸入會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboTMPML.SetFocus
        Exit Function
    End If
    
    
    If Chk_MLClass(cboTMPML, "A", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboTMPML.SetFocus
        lblDspTmpMlDesc = ""
       Exit Function
    End If
    
    lblDspTmpMlDesc = wsDesc
    
    Chk_cboTmpML = True
    
End Function


Private Sub txtRmk_GotFocus()

        
        FocusMe txtRmk

End Sub

Private Sub txtRmk_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtRmk, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        cboCusCode.SetFocus
        
        
    End If
End Sub

Private Sub txtRmk_LostFocus()
        
        FocusMe txtRmk, True

End Sub




Private Function Chk_DocNo(ByVal InDocNo As String, ByRef OutStatus As String, ByRef OutUpdFlg As String, ByRef OutDocDate As String, ByRef OutPgmNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    OutStatus = ""
    OutUpdFlg = ""
    Chk_DocNo = False
    
    wsSQL = "SELECT ARCQCHQDATE, ARCQSTATUS, ARCQUPDFLG, ARCQPGMNO FROM ARCHEQUE "
    wsSQL = wsSQL & " WHERE ARCQCHQNO = '" & Set_Quote(InDocNo) & "' "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount <= 0 Then
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    OutDocDate = ReadRs(rsRcd, "ARCQCHQDATE")
    OutStatus = ReadRs(rsRcd, "ARCQSTATUS")
    OutUpdFlg = ReadRs(rsRcd, "ARCQUPDFLG")
    OutPgmNo = ReadRs(rsRcd, "ARCQPGMNO")
    rsRcd.Close
    Set rsRcd = Nothing
    Chk_DocNo = True
    
    
    

End Function



Private Function Chk_CusCode(ByVal InCusNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT CusID, CusName, CusTel, CusFax FROM mstCustomer WHERE CusCode = '" & Set_Quote(InCusNo) & "' "
    wsSQL = wsSQL & "And CusStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "CusID")
        OutName = ReadRs(rsRcd, "CusName")
        OutTel = ReadRs(rsRcd, "CusTel")
        OutFax = ReadRs(rsRcd, "CusFax")
        Chk_CusCode = True
        
    Else
    
        OutID = 0
        OutName = ""
        OutTel = ""
        OutFax = ""
        Chk_CusCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Private Function Chk_txtChqAmtOrg(inAmt As String) As Integer
    
    Chk_txtChqAmtOrg = False
    
   
    If To_Value(inAmt) <= 0 Then
        gsMsg = "Cheque Amount Cannot equal to 0!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtChqAmtOrg.SetFocus
        Exit Function
    End If
    
    
    If To_Value(inAmt) > gsMaxVal Then
        gsMsg = "數量太大!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtChqAmtOrg.SetFocus
        Exit Function
    End If
    
    Chk_txtChqAmtOrg = True

End Function

