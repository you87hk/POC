VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmITMLST 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmITMLST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10061.66
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   1812
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9000
      OleObjectBlob   =   "frmITMLST.frx":0442
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   9735
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   1890
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   9600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":341F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":3CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":414B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":459D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":48B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":4D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":578F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":5BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":64BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":67E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":6C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITMLST.frx":6F55
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "frmITMLST.frx":7271
      TabIndex        =   1
      Top             =   1320
      Width           =   9855
   End
End
Attribute VB_Name = "frmITMLST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB

Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Dim waInvDoc As New XArrayDB
Private wbErr As Boolean
Private wlLineNo As Long
Private wsBaseCurCd As String


Private wiExit As Boolean
Private wiUpdate As Boolean

Private wsFormCaption As String
Private wsFormID As String


Private Const tcGo = "Go"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"

Private Const XSEL = 0
Private Const XITMTYPE = 1
Private Const XITMCODE = 2
Private Const XITMCLS = 3
Private Const XITMNAME = 4
Private Const XUNITPRICE = 5
Private Const XITMID = 6

        
Private Const LINENO = 0
Private Const ITMTYPE = 1
Private Const ITMCODE = 2
Private Const ITMNAME = 3
Private Const LOTNO = 4
Private Const WHSCODE = 5
Private Const SOH = 6
Private Const LOTTO = 7
Private Const PRICE = 8
Private Const QTY = 9
Private Const NET = 10
Private Const ITMID = 11
Private Const SOID = 12



Public Property Get InvDoc() As XArrayDB
    Set InvDoc = waInvDoc
End Property

Public Property Let InvDoc(inInvDoc As XArrayDB)
    Set waInvDoc = inInvDoc
End Property


Public Property Get inLineNo() As Long
     inLineNo = wlLineNo
End Property

Public Property Let inLineNo(inLine As Long)
     wlLineNo = inLine
End Property



Private Sub cboDocNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, CUSNAME "
    wsSQL = wsSQL & " FROM SOASOHD, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS <> '2' "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    
    Call Ini_Combo(3, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLITMLIST", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(cboDocNoFr, 15, KeyAscii, True)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If chk_cboDocNoFr = False Then Exit Sub
            
            If LoadRecord = True Then
                tblDetail.SetFocus
            End If
       
    End If
End Sub

Private Sub cboDocNoFr_LostFocus()
    FocusMe cboDocNoFr, True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF10
           Call cmdOK
            
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF7
            Call LoadRecord
        
        Case vbKeyF5
            Call cmdSelect(1)

           
        Case vbKeyF6
            Call cmdSelect(0)

       
    End Select
End Sub

Private Sub Form_Load()
    
    
  MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Grid
    Ini_Scr

    
   MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCancel()
    
    
  MousePointer = vbHourglass
  
    Ini_Scr
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub cmdOK()
    
    
  MousePointer = vbHourglass
  Unload Me
  MousePointer = vbDefault
    
    
End Sub
Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, XSEL, XITMID
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
   '         Case "ComboBox"
   '             MyControl.Clear
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

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    wiExit = False
    wiUpdate = True
    
    
    cboDocNoFr.Text = ""


End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If wiExit = False Then
     
       Cancel = True
       If wiUpdate = True Then Call UpdateRecord
       wiExit = True
       Me.Hide
       Exit Sub
    End If
    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set waInvDoc = Nothing
    Set frmITMLST = Nothing
    
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    
    wsFormID = "ITMLST"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")

    
    With tblDetail
        .Columns(XSEL).Caption = Get_Caption(waScrItm, "XSEL")
        .Columns(XITMCODE).Caption = Get_Caption(waScrItm, "XITMCODE")
        .Columns(XITMTYPE).Caption = Get_Caption(waScrItm, "XITMTYPE")
        .Columns(XITMCLS).Caption = Get_Caption(waScrItm, "XITMCLS")
        .Columns(XITMNAME).Caption = Get_Caption(waScrItm, "XITMNAME")
        .Columns(XUNITPRICE).Caption = Get_Caption(waScrItm, "XUNITPRICE")
        .Columns(XITMID).Caption = Get_Caption(waScrItm, "XITMID")
    
    End With
    
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F10)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F5)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F6)"
    
    

End Sub



Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .Update
    End With
End Sub



Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
       
        Case vbKeyReturn
            Select Case .Col

                Case XUNITPRICE
                    KeyCode = vbDefault
                    .Col = XSEL
                Case XITMTYPE, XITMCODE, XITMCLS, XITMNAME
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
              If .Col <> XSEL Then
                    .Col = .Col - 1
              End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> XUNITPRICE Then
                    .Col = .Col + 1
            End If
            
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub






Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Go"
            Call cmdOK
            
        Case "Cancel"
        
           Call cmdCancel
            
        Case "Exit"
           wiUpdate = False
           Unload Me
            
        Case "Refresh"
            Call LoadRecord
            
        Case tcSAll
        
           Call cmdSelect(1)
        
        Case tcDAll
        
           Call cmdSelect(0)
            
    End Select
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If

End Sub


Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = False
        .MultipleLines = 0
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
     '   .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = XSEL To XITMID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case XSEL
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).Locked = False
                Case XITMTYPE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 10
                Case XITMCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 30
                Case XITMCLS
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 1
                Case XITMNAME
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 60
                Case XUNITPRICE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                Case XITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                    
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim lstUPrice As Double
    Dim lstExcR As Double
    Dim lstCurr As String
    

    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    wsSQL = "SELECT SODTID DTID, ITMCODE, ITMITMTYPECODE, ITMCLASS, SODTITEMDESC ITMNAME, "
    wsSQL = wsSQL & "SODTUPRICE UPRICE "
    wsSQL = wsSQL & "FROM  SOASOHD, SOASODT, MSTITEM "
    wsSQL = wsSQL & "WHERE SOHDDOCNO LIKE '%" & Set_Quote(Trim(cboDocNoFr)) & "%'"
    wsSQL = wsSQL & "AND SOHDDOCID = SODTDOCID "
    wsSQL = wsSQL & "AND SODTITEMID = ITMID "
    wsSQL = wsSQL & "ORDER BY ITMITMTYPECODE, ITMCODE "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
        
    
    With waResult
    .ReDim 0, -1, XSEL, XITMID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
        
        
     .AppendRows
        waResult(.UpperBound(1), XITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
        waResult(.UpperBound(1), XITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), XITMCLS) = ReadRs(rsRcd, "ITMCLASS")
        waResult(.UpperBound(1), XITMNAME) = ReadRs(rsRcd, "ITMNAME")
        waResult(.UpperBound(1), XUNITPRICE) = Format(ReadRs(rsRcd, "UPRICE"), gsUprFmt)
        waResult(.UpperBound(1), XITMID) = ReadRs(rsRcd, "DTID")
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    LoadRecord = True
    Me.MousePointer = vbNormal
    
    
End Function

Private Function UpdateRecord() As Boolean
    Dim wiCtr As Long
    
    UpdateRecord = False
    
    
    If waResult.UpperBound(1) >= 0 Then
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, XSEL)) = "-1" Then
                
                Call InsRecord(waResult(wiCtr, XITMID))
                wlLineNo = wlLineNo + 1
 
            End If
        Next
    End If
    
    
    
    UpdateRecord = True
    


    
End Function


Private Function InsRecord(ByVal inItmID As String) As Boolean
    Dim wiCtr As Long
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim lstUPrice As Double
    
    
    Dim lstExcR As String
    Dim lstCurr As String
    
    
    
    Dim wiRow As Long
    
    InsRecord = False
    
    
    wsSQL = "SELECT SODTITEMID ITMID, ITMCODE, ITMITMTYPECODE, ITMBARCODE, ITMCURR, SODTITEMDESC ITMNAME, "
    wsSQL = wsSQL & "ITMUNITPRICE UPRICE, SODTQTY QTY "
    wsSQL = wsSQL & "FROM  SOASODT, MSTITEM "
    wsSQL = wsSQL & "WHERE SODTID = " & To_Value(inItmID)
    wsSQL = wsSQL & "AND SODTITEMID = ITMID "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    lstUPrice = To_Value(ReadRs(rsRcd, "UPRICE"))
    lstCurr = ReadRs(rsRcd, "ITMCURR")
        
    If UCase(Trim(lstCurr)) <> UCase(wsBaseCurCd) Then
    
        Call getExcRate(lstCurr, gsSystemDate, lstExcR, "")
        lstUPrice = Format(lstUPrice * To_Value(lstExcR), gsAmtFmt)
        
    End If
        

  
        
    With waInvDoc
            .AppendRows
             waInvDoc(.UpperBound(1), LINENO) = wlLineNo
             waInvDoc(.UpperBound(1), ITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
             waInvDoc(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
           '  waInvDoc(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waInvDoc(.UpperBound(1), ITMNAME) = ReadRs(rsRcd, "ITMNAME")
             waInvDoc(.UpperBound(1), LOTNO) = ""
             waInvDoc(.UpperBound(1), WHSCODE) = ""
         '    waInvDoc(.UpperBound(1), PUBLISHER) = ""
             waInvDoc(.UpperBound(1), QTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
             waInvDoc(.UpperBound(1), PRICE) = Format(lstUPrice, gsAmtFmt)
             waInvDoc(.UpperBound(1), NET) = Format(lstUPrice * To_Value(ReadRs(rsRcd, "QTY")), gsAmtFmt)
             waInvDoc(.UpperBound(1), ITMID) = ReadRs(rsRcd, "ITMID")
             waInvDoc(.UpperBound(1), SOID) = "0"
             
    End With
    
      rsRcd.Close
    Set rsRcd = Nothing
  

  
    InsRecord = True
    


    
End Function


Private Function chk_cboDocNoFr() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
 
    
    chk_cboDocNoFr = False
    
     wsSQL = "SELECT * FROM SOASOHD "
     wsSQL = wsSQL & "WHERE SOHDDOCNO = '" & Set_Quote(cboDocNoFr.Text) & "'"
     
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    chk_cboDocNoFr = True
    
End Function
Private Sub cmdSelect(ByVal wiSelect As Integer)
    Dim wiCtr As Long
    
    Me.MousePointer = vbHourglass
    
    
     
    With waResult
    For wiCtr = 0 To .UpperBound(1)
        waResult(wiCtr, XSEL) = IIf(wiSelect = 1, "-1", "0")
    Next wiCtr
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    Me.MousePointer = vbNormal
    
End Sub
