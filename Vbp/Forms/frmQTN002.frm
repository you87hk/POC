VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmQTN002 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmQTN002.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9000
      OleObjectBlob   =   "frmQTN002.frx":0442
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   4575
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":341F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":3CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":414B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":459D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":48B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":4D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":578F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":5BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":64BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN002.frx":67E5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Object.Visible         =   0   'False
            Key             =   "BOM"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6855
      Left            =   120
      OleObjectBlob   =   "frmQTN002.frx":6C39
      TabIndex        =   2
      Top             =   480
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
Attribute VB_Name = "frmQTN002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control

Private waPopUpSub As New XArrayDB

Dim waInvDoc As New XArrayDB
Dim wlHLineNo As Long
Dim wlLineNo As Long
Dim wsHCurr As String
Dim wdHExcr As Double



Private wbErr As Boolean

Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String

Private Const tcGo = "Go"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcBOM = "BOM"

Private Const SLINENO = 0
Private Const SLN = 1
Private Const SINDENT = 2
Private Const SITMTYPE = 3
Private Const SITMCODE = 4
Private Const SVDRCODE = 5
Private Const SITMNAME = 6
Private Const SQTY = 7
Private Const SUCST = 8
Private Const SCST = 9
Private Const SUNITPRICE = 10
Private Const SDISPER = 11
Private Const SAMT = 12
Private Const SNET = 13
Private Const SVDRID = 14
Private Const SITMID = 15



Public Property Get InvDoc() As XArrayDB
    Set InvDoc = waInvDoc
End Property

Public Property Let InvDoc(inInvDoc As XArrayDB)
    Set waInvDoc = inInvDoc
End Property

Public Property Let inLineNo(inLine As Long)
    wlHLineNo = inLine
End Property

Public Property Let InCurr(InCurCd As String)
    wsHCurr = InCurCd
End Property

Public Property Let InExcr(inExcRate As Double)
    wdHExcr = inExcRate
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
           Call cmdOK
            
        Case vbKeyF3
           Call cmdCancel
            
        Case vbKeyF12
            Me.Hide
        
        Case vbKeyF5
            Call LoadRecord
        
        Case vbKeyF9
            Call cmdBOM
            
       ' Case vbKeyEscape
       '     Unload Me
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
    
    waResult.ReDim 0, -1, SLINENO, SITMID
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
    
    
   Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If wiExit = False Then
        If waResult.UpperBound(1) >= 0 Then
            With tblDetail
                .Update
                If Chk_GrdRow(.FirstRow + .Row) = False Then
                    .SetFocus
                    Exit Sub
                End If
            End With
        End If
    
       Cancel = True
       Call UpdateRecord
       wiExit = True
       Me.Hide
       Exit Sub
    End If
    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set waInvDoc = Nothing
    Set waPopUpSub = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "QTN002"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP_A", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    With tblDetail
        .Columns(SLINENO).Caption = Get_Caption(waScrItm, "SLINENO")
        .Columns(SINDENT).Caption = Get_Caption(waScrItm, "SINDENT")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SVDRCODE).Caption = Get_Caption(waScrItm, "SVDRCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SUNITPRICE).Caption = Get_Caption(waScrItm, "SUNITPRICE")
        .Columns(SUCST).Caption = Get_Caption(waScrItm, "SUCST")
        .Columns(SDISPER).Caption = Get_Caption(waScrItm, "SDISPER")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SAMT).Caption = Get_Caption(waScrItm, "SAMT")
        .Columns(SNET).Caption = Get_Caption(waScrItm, "SNET")
        .Columns(SCST).Caption = Get_Caption(waScrItm, "SCST")
        
    End With
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F2)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcBOM).ToolTipText = Get_Caption(waScrToolTip, tcBOM) & "(F9)"
    
    

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

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case tcGo
            Call cmdOK
            
        Case tcCancel
        
           Call cmdCancel
            
        Case tcBOM
            Call cmdBOM
            
        Case tcExit
            Me.Hide
            
        Case tcRefresh
            Call LoadRecord
            
    End Select
End Sub



Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case SITMTYPE
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
              Case SITMTYPE
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


Private Function LoadRecord() As Boolean
    Dim wiRow As Long
    
    LoadRecord = False

    wlLineNo = 1
    
    If waInvDoc.UpperBound(1) >= 0 Then
    With waResult
    .ReDim 0, -1, SLINENO, SITMID
         
    For wiRow = 0 To waInvDoc.UpperBound(1)
         
         If waInvDoc(wiRow, SLN) = wlHLineNo Then
             .AppendRows
             waResult(.UpperBound(1), SLINENO) = waInvDoc(wiRow, SLINENO)
             waResult(.UpperBound(1), SLN) = wlHLineNo
             waResult(.UpperBound(1), SINDENT) = waInvDoc(wiRow, SINDENT)
             waResult(.UpperBound(1), SITMTYPE) = waInvDoc(wiRow, SITMTYPE)
             waResult(.UpperBound(1), SITMCODE) = waInvDoc(wiRow, SITMCODE)
             waResult(.UpperBound(1), SVDRCODE) = waInvDoc(wiRow, SVDRCODE)
             waResult(.UpperBound(1), SITMNAME) = waInvDoc(wiRow, SITMNAME)
             waResult(.UpperBound(1), SUNITPRICE) = Format(waInvDoc(wiRow, SUNITPRICE), gsAmtFmt)
             waResult(.UpperBound(1), SUCST) = Format(waInvDoc(wiRow, SUCST), gsAmtFmt)
             
             waResult(.UpperBound(1), SDISPER) = Format(waInvDoc(wiRow, SDISPER), gsAmtFmt)
             waResult(.UpperBound(1), SQTY) = Format(waInvDoc(wiRow, SQTY), gsQtyFmt)
             waResult(.UpperBound(1), SAMT) = Format(waInvDoc(wiRow, SAMT), gsAmtFmt)
             waResult(.UpperBound(1), SNET) = Format(waInvDoc(wiRow, SNET), gsAmtFmt)
             waResult(.UpperBound(1), SCST) = Format(waInvDoc(wiRow, SCST), gsAmtFmt)
             
             waResult(.UpperBound(1), SITMID) = To_Value(waInvDoc(wiRow, SITMID))
             waResult(.UpperBound(1), SVDRID) = To_Value(waInvDoc(wiRow, SVDRID))
            wlLineNo = wlLineNo + 1
         End If
         
 
    
    Next wiRow
    End With
    
    tblDetail.ReBind
    'tblDetail.FirstRow = 0
    tblDetail.Bookmark = 0
    
    End If
   
    
    
    LoadRecord = True
    
End Function

Private Function UpdateRecord() As Boolean
    Dim wiCtr As Long
    
    UpdateRecord = False

    With waInvDoc
  
    If waInvDoc.UpperBound(1) >= 0 Then
    '  .ReDim 0, waInvDoc.UpperBound(1), STYPE, SSTATUS
    For wiCtr = 0 To waInvDoc.UpperBound(1)
        If waInvDoc(wiCtr, SLN) = wlHLineNo Then
            waInvDoc(wiCtr, SLN) = "0"
        End If
    Next wiCtr
    End If
         
    If waResult.UpperBound(1) >= 0 Then
         
    For wiCtr = 0 To waResult.UpperBound(1)
         If Trim(waResult(wiCtr, SLINENO)) <> "" Then
             .AppendRows
             waInvDoc(.UpperBound(1), SLINENO) = waResult(wiCtr, SLINENO)
             waInvDoc(.UpperBound(1), SINDENT) = waResult(wiCtr, SINDENT)
             waInvDoc(.UpperBound(1), SITMTYPE) = waResult(wiCtr, SITMTYPE)
             waInvDoc(.UpperBound(1), SITMCODE) = waResult(wiCtr, SITMCODE)
             waInvDoc(.UpperBound(1), SVDRCODE) = waResult(wiCtr, SVDRCODE)
             waInvDoc(.UpperBound(1), SITMNAME) = waResult(wiCtr, SITMNAME)
             waInvDoc(.UpperBound(1), SUNITPRICE) = Format(waResult(wiCtr, SUNITPRICE), gsAmtFmt)
             waInvDoc(.UpperBound(1), SUCST) = Format(waResult(wiCtr, SUCST), gsAmtFmt)
             
             waInvDoc(.UpperBound(1), SDISPER) = Format(waResult(wiCtr, SDISPER), gsAmtFmt)
             
             waInvDoc(.UpperBound(1), SQTY) = Format(waResult(wiCtr, SQTY), gsQtyFmt)
             waInvDoc(.UpperBound(1), SAMT) = Format(waResult(wiCtr, SAMT), gsAmtFmt)
             waInvDoc(.UpperBound(1), SNET) = Format(waResult(wiCtr, SNET), gsAmtFmt)
             waInvDoc(.UpperBound(1), SCST) = Format(waResult(wiCtr, SCST), gsAmtFmt)
             
             waInvDoc(.UpperBound(1), SLN) = wlHLineNo
             waInvDoc(.UpperBound(1), SITMID) = To_Value(waResult(wiCtr, SITMID))
             waInvDoc(.UpperBound(1), SVDRID) = To_Value(waResult(wiCtr, SVDRID))
         End If
    Next wiCtr
    End If
   
    End With
    
    UpdateRecord = True
    
End Function


Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(SITMTYPE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, SLINENO)) = "" And _
                   Trim(waResult(inRow, SITMTYPE)) = "" And _
                   Trim(waResult(inRow, SITMCODE)) = "" And _
                   Trim(waResult(inRow, SVDRCODE)) = "" And _
                   Trim(waResult(inRow, SITMNAME)) = "" And _
                   Trim(waResult(inRow, SQTY)) = "" And _
                   Trim(waResult(inRow, SUNITPRICE)) = "" And _
                   Trim(waResult(inRow, SUCST)) = "" And _
                   Trim(waResult(inRow, SNET)) = "" And _
                   Trim(waResult(inRow, SCST)) = "" And _
                   Trim(waResult(inRow, SDISPER)) = "" And _
                   Trim(waResult(inRow, SAMT)) = "" And _
                   Trim(waResult(inRow, SITMID)) = "" And _
                   Trim(waResult(inRow, SVDRID)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
End Function

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
Dim sTemp As String
   
    With tblDetail
        sTemp = .Columns(ColIndex)
        .Update
    End With


 '   If ColIndex = ItmCode Then
 '       Call LoadBookGroup(sTemp)
 '   End If
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsITMID As String
Dim wsITMCODE As String
Dim wsVdrID As String
Dim wsVdrCODE As String
Dim wsITMTYPE As String
Dim wsITMNAME As String
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
        
            Case SITMTYPE
                If Chk_grdItmType(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                
                If Trim(.Columns(SITMCODE).Text) <> "" Then
                
                If Chk_grdITMCODE(.Columns(SITMCODE).Text, .Columns(SITMTYPE).Text, wsITMID, wsITMCODE, wsITMNAME, wsVdrID, wsVdrCODE) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(SLINENO).Text = wlLineNo
                .Columns(SITMID).Text = wsITMID
                .Columns(SVDRID).Text = wsVdrID
                .Columns(SITMNAME).Text = wsITMNAME
                .Columns(SVDRCODE).Text = wsVdrCODE
                .Columns(SQTY).Text = "0"
                .Columns(SUNITPRICE).Text = get_ItemSalePrice(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                .Columns(SUCST).Text = get_ItemVdrCost(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                
                .Columns(SLN).Text = wlHLineNo
                .Columns(SDISPER).Text = Format("0", gsAmtFmt)
                
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(SITMCODE).Text) <> wsITMCODE Then
                    .Columns(SITMCODE).Text = wsITMCODE
                End If
                
                End If
                
                
            Case SITMCODE
               ' If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                '    GoTo Tbl_BeforeColUpdate_Err
                'End If
                
                If Chk_grdITMCODE(.Columns(SITMCODE).Text, .Columns(SITMTYPE).Text, wsITMID, wsITMCODE, wsITMNAME, wsVdrID, wsVdrCODE) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(SLINENO).Text = wlLineNo
                .Columns(SITMID).Text = wsITMID
                .Columns(SVDRID).Text = wsVdrID
                .Columns(SITMNAME).Text = wsITMNAME
                .Columns(SVDRCODE).Text = wsVdrCODE
                .Columns(SQTY).Text = "0"
                .Columns(SUNITPRICE).Text = get_ItemSalePrice(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                .Columns(SUCST).Text = get_ItemVdrCost(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                
                .Columns(SDISPER).Text = Format("0", gsAmtFmt)
                
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(SITMCODE).Text) <> wsITMCODE Then
                    .Columns(SITMCODE).Text = wsITMCODE
                End If
                
            Case SVDRCODE
                
                If Chk_grdVdrCode(.Columns(ColIndex).Text, wsVdrID) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
            
                .Columns(SVDRID).Text = wsVdrID
                .Columns(SUNITPRICE).Text = get_ItemSalePrice(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                .Columns(SUCST).Text = get_ItemVdrCost(.Columns(SITMID).Text, .Columns(SVDRID).Text, wsHCurr, wdHExcr)
                
                
           Case SUNITPRICE, SQTY, SDISPER
                
                If ColIndex = SUNITPRICE Then
                
                If Chk_grdUNITPRICE(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                End If
                
                If ColIndex = SQTY Then
                
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                End If
                
                If ColIndex = SDISPER Then
                
                If Chk_grdDisPer(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                End If
                
                If Trim(.Columns(SUNITPRICE)) <> "" And Trim(.Columns(SQTY)) <> "" And Trim(.Columns(SDISPER)) <> "" Then
                    .Columns(SAMT).Text = Format(To_Value(.Columns(SUNITPRICE)) * To_Value(.Columns(SQTY)), gsAmtFmt)
                End If
                
                If Trim(.Columns(SUNITPRICE)) <> "" And Trim(.Columns(SQTY)) <> "" And Trim(.Columns(SDISPER)) <> "" Then
                    .Columns(SNET).Text = Format(To_Value(.Columns(SUNITPRICE)) * To_Value(.Columns(SQTY)) * (1 - To_Value(.Columns(SDISPER))), gsAmtFmt)
                End If
                
                If Trim(.Columns(SUCST)) <> "" And Trim(.Columns(SQTY)) <> "" Then
                    .Columns(SCST).Text = Format(To_Value(.Columns(SUCST)) * To_Value(.Columns(SQTY)), gsAmtFmt)
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
            Case SITMCODE
                
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMITMTYPECODE, ITMENGNAME ITNAME, STR(ITMDEFAULTPRICE,13,2) FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(SITMCODE).Text) & "%' "
                    wsSql = wsSql & " AND ITMITMTYPECODE =  '" & Set_Quote(.Columns(SITMTYPE).Text) & "' "
                    
                   ' If waResult.UpperBound(1) > -1 Then
                   '     wsSql = wsSql & " AND ITMCODE NOT IN ( "
                   '     For wiCtr = 0 To waResult.UpperBound(1)
                   '         wsSql = wsSql & " '" & waResult(wiCtr, SITMCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                   '     Next
                   ' End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMITMTYPECODE, ITMCHINAME ITNAME, STR(ITMDEFAULTPRICE,13,2) FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(SITMCODE).Text) & "%' "
                    wsSql = wsSql & " AND ITMITMTYPECODE =  '" & Set_Quote(.Columns(SITMTYPE).Text) & "' "
                    
                  '  If waResult.UpperBound(1) > -1 Then
                  '      wsSql = wsSql & " AND ITMCODE NOT IN ( "
                  '      For wiCtr = 0 To waResult.UpperBound(1)
                  '          wsSql = wsSql & " '" & waResult(wiCtr, SITMCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                  '      Next
                  '  End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(4, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case SITMTYPE
            
              
            If gsLangID = "2" Then
    
            wsSql = "SELECT ITMTYPECODE, ITMTYPECHIDESC "
            wsSql = wsSql & " FROM MSTITEMTYPE "
            wsSql = wsSql & " WHERE ITMTYPECODE LIKE '%" & Set_Quote(.Columns(SITMTYPE).Text) & "%' "
            wsSql = wsSql & " AND ITMTYPESTATUS  <> '2' "
            wsSql = wsSql & " ORDER BY ITMTYPECODE "
    
            Else
    
            wsSql = "SELECT ITMTYPECODE, ITMTYPEENGDESC "
            wsSql = wsSql & " FROM MSTITEMTYPE "
            wsSql = wsSql & " WHERE ITMTYPECODE LIKE '%" & Set_Quote(.Columns(SITMTYPE).Text) & "%' "
            wsSql = wsSql & " AND ITMTYPESTATUS  <> '2' "
            wsSql = wsSql & " ORDER BY ITMTYPECODE "
    
            End If

                
            '    wsSql = "SELECT ITMITMTYPECODE, ITMCODE FROM mstItem "
            '    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMITMTYPECODE LIKE '%" & Set_Quote(.Columns(SITMTYPE).Text) & "%' "
            '    wsSql = wsSql & " ORDER BY ITMITMTYPECODE "
               
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLITMTYPE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case SVDRCODE
                
                wsSql = "SELECT VDRCODE, VDRNAME FROM mstVENDOR "
                wsSql = wsSql & " WHERE VDRSTATUS <> '2' AND VDRCODE LIKE '%" & Set_Quote(.Columns(SVDRCODE).Text) & "%' "
                wsSql = wsSql & " ORDER BY VDRCODE "
               
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLVDRCODE", Me.Width, Me.Height)
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
                Case SINDENT
                    KeyCode = vbDefault
                    .Col = SITMTYPE
                Case SITMTYPE
                    KeyCode = vbDefault
                    .Col = SITMCODE
                Case SITMCODE
                    KeyCode = vbDefault
                    .Col = SQTY
                Case SVDRCODE
                    KeyCode = vbDefault
                    .Col = SITMNAME
                Case SITMNAME
                    KeyCode = vbDefault
                    .Col = SQTY
               Case SQTY
                    KeyCode = vbDefault
                    .Col = SUNITPRICE
                Case SUNITPRICE
                    KeyCode = vbDefault
                    .Col = SDISPER
                Case SDISPER
                    KeyCode = vbKeyDown
                    .Col = SITMTYPE
                
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case SITMTYPE
                    .Col = SINDENT
                Case SITMCODE
                    .Col = SITMTYPE
                Case SVDRCODE
                    .Col = SITMCODE
                Case SITMNAME
                    .Col = SVDRCODE
                Case SQTY
                    .Col = SITMNAME
                Case SUCST
                    .Col = SQTY
                Case SCST
                .Col = SUCST
                Case SUNITPRICE
                .Col = SCST
                Case SDISPER
                .Col = SUNITPRICE
                Case SAMT
                    .Col = SDISPER
                Case SNET
                    .Col = SAMT
                    
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case SINDENT
                    .Col = SITMTYPE
                Case SITMTYPE
                    .Col = SITMCODE
                Case SITMCODE
                    .Col = SVDRCODE
                Case SVDRCODE
                    .Col = SITMNAME
                Case SITMNAME
                    .Col = SQTY
                Case SQTY
                    .Col = SUCST
                Case SUCST
                    .Col = SCST
                Case SCST
                    .Col = SUNITPRICE
                Case SUNITPRICE
                    .Col = SDISPER
                Case SDISPER
                    .Col = SAMT
                Case SAMT
                    .Col = SNET
            End Select
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    Select Case tblDetail.Col
        Case SUNITPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
        
        Case SDISPER
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
        Case SQTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
    
    End Select
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = SITMTYPE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case SITMCODE
                    Call Chk_grdITMCODE(.Columns(SITMCODE).Text, .Columns(SITMTYPE).Text, "", "", "", "", "")
               Case SITMTYPE
                    Call Chk_grdItmType(.Columns(SITMTYPE).Text)
               Case SVDRCODE
                    Call Chk_grdVdrCode(.Columns(SVDRCODE).Text, "")
               Case SDISPER
                    Call Chk_grdDisPer(.Columns(SDISPER).Text)
                'Case QTY
                '    Call Chk_grdQTY(.Columns(QTY).Text)
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdITMCODE(ByVal inAccNo As String, ByVal inITMTYPE As String, ByRef outAccID As String, ByRef outAccNo As String, ByRef OutName As String, ByRef OutVdrID As String, ByRef OutVdrCode As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    Dim wlVdrID As Long
    
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入書號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
        Exit Function
    End If
    
    wlVdrID = 0
    
    wsSql = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMPVDRID, ITMBOTTOMPRICE, ITMCURR FROM MSTITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' AND ITMITMTYPECODE = '" & Set_Quote(inITMTYPE) & "' "
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        wlVdrID = To_Value(ReadRs(rsDes, "ITMPVDRID"))
        
        Chk_grdITMCODE = True
    Else
        outAccID = ""
        OutName = ""
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
        rsDes.Close
        Set rsDes = Nothing
        Exit Function
    End If
    
    rsDes.Close
    Set rsDes = Nothing
        
    If wlVdrID = 0 Then
    
    wsSql = "SELECT VdrItemVdrID VID, VdrCode "
    wsSql = wsSql & " FROM MstVdrItem, MstVendor "
    wsSql = wsSql & " WHERE VdrItemItmID = " & outAccID
    wsSql = wsSql & " AND VdrItemStatus = '1' "
    wsSql = wsSql & " AND VdrItemVdrID = VdrID "
    wsSql = wsSql & " Order By VdrItemCostl "
    
    Else
    
    wsSql = "SELECT VdrID VID, VdrCode "
    wsSql = wsSql & " FROM MstVendor "
    wsSql = wsSql & " WHERE VdrID = " & wlVdrID
    
    End If
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        
        OutVdrID = ReadRs(rsDes, "VID")
        OutVdrCode = ReadRs(rsDes, "VdrCode")

    Else
    
        OutVdrID = ""
        OutVdrCode = ""

    End If
    
    rsDes.Close
    Set rsDes = Nothing
    
    
    
End Function


Private Function Chk_grdItmType(inAccNo As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdItmType = False
        Exit Function
    End If
    
    
    wsSql = "SELECT * FROM MSTITEMTYPE"
    wsSql = wsSql & " WHERE ITMTYPECODE = '" & Set_Quote(inAccNo) & "'"
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        Chk_grdItmType = True
    Else
        gsMsg = "沒有此分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdItmType = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
End Function
Private Function Chk_grdVdrCode(ByVal inAccNo As String, ByRef outAccID As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdVdrCode = False
        Exit Function
    End If
    
    
    wsSql = "SELECT VDRID FROM MSTVENDOR "
    wsSql = wsSql & " WHERE VdrCode = '" & Set_Quote(inAccNo) & "'"
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
    
     outAccID = ReadRs(rsDes, "VdrID")
     Chk_grdVdrCode = True
        
    Else
        outAccID = ""
        gsMsg = "沒有此分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdVdrCode = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
    
End Function
Private Function Chk_grdQty(inQty As String) As Boolean
    Chk_grdQty = False
    
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
    
    Chk_grdQty = True
End Function

Private Function Chk_grdUNITPRICE(inUnitPrice As String) As Boolean
    Chk_grdUNITPRICE = False
    
    If Trim(inUnitPrice) = "" Then
        gsMsg = "沒有輸入價格!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    If To_Value(inUnitPrice) < 0 Then
        gsMsg = "價格不可小於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdUNITPRICE = True
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
        
        If Chk_grdItmType(waResult(LastRow, SITMTYPE)) = False Then
            .Col = SITMTYPE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdITMCODE(waResult(LastRow, SITMCODE), waResult(LastRow, SITMTYPE), "", "", "", "", "") = False Then
            .Col = SITMCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdVdrCode(waResult(LastRow, SVDRCODE), "") = False Then
            .Col = SVDRCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdUNITPRICE(waResult(LastRow, SUNITPRICE)) = False Then
            .Col = SUNITPRICE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdDisPer(waResult(LastRow, SDISPER)) = False Then
            .Col = SDISPER
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdQty(waResult(LastRow, SQTY)) = False Then
            .Col = SQTY
            .Row = LastRow
            Exit Function
        End If
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function


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
        
        For wiCtr = SLINENO To SITMID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SLINENO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Locked = True
                Case SLN
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Visible = False
                Case SINDENT
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 2
                Case SITMCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                Case SITMTYPE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                Case SVDRCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                Case SITMNAME
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 60

                Case SUNITPRICE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                 Case SUCST
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case SDISPER
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 6
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SQTY
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Locked = False
                Case SAMT
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SNET
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SCST
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                
                Case SITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case SVDRID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                    
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoDup = False
    
        wsCurRec = tblDetail.Columns(SITMCODE)
 '       wsCurRecLn = tblDetail.Columns(wsWhsCode)
 
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, SITMCODE) Then
                  gsMsg = "重覆於第 " & waResult(wlCtr, SLINENO) & " 行!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function


Private Sub cmdBOM()
        
      '  With frmITMLST
      '      .inCurr = wsHCurr
      '      .inExcr = wdHExcr
      '      .InvDoc = waResult
      '      .InvItem = waItem
      '      .inLineNo = wlLineNo
      '      .Show vbModal
      '      waResult.ReDim 0, .InvDoc.UpperBound(1), GLINENO, GDESC4
      '      waItem.ReDim 0, .InvItem.UpperBound(1), SLINENO, SITMID
      '      Set waResult = .InvDoc
      '      Set waItem = .InvItem
      '      wlLineNo = .inLineNo
      '  End With
        
      '  Unload frmITMLST
      '  tblDetail.ReBind
      '  tblDetail.Bookmark = 0
        
        'Call Calc_Total
        'Call cmdCstRefresh
        
        
End Sub
Private Function Chk_grdDisPer(inCode As String) As Boolean
    
    Chk_grdDisPer = True
    

    If To_Value(inCode) < 0 Then
        gsMsg = "單價必需大為零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdDisPer = False
        Exit Function
    End If
    
End Function
