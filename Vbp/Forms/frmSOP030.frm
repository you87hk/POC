VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSOP030 
   Caption         =   "SOP030"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   9195
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   9000
      OleObjectBlob   =   "frmSOP030.frx":0000
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2790
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2040
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
            Picture         =   "frmSOP030.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP030.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2790
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   3015
      Left            =   2760
      OleObjectBlob   =   "frmSOP030.frx":63A3
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label lblPrice 
      Caption         =   "PRICE"
      Height          =   225
      Left            =   870
      TabIndex        =   7
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "PRDTO"
      Height          =   225
      Left            =   4920
      TabIndex        =   6
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "PRDFR"
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   1230
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   4
      Top             =   760
      Width           =   1860
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
Attribute VB_Name = "frmSOP030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Dim wgsTitle As String
Private wsFormCaption As String

Private waPopUpSub As New XArrayDB
Private Const FRPRICE = 0
Private Const TOPRICE = 1
Private Const GDummy = 2

Private waResult As New XArrayDB
Private wiAction As Integer
Private wbUpdate As Boolean
Private wbErr As Boolean

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String

Private Sub cmdCancel()
    Ini_Scr
    medPrdFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim adcmdSave As New ADODB.Command
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim sFrPrd As String
    Dim sToPrd As String
    Dim wiCtr As Integer
    Dim iPos As Integer
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(1)
    'wsSelection(1) = medPrdFr.Caption & " " & Set_Quote(cboCusNoFr.Text)
    
    wsDteTim = Now
    
    iPos = InStr(1, medPrdFr, "/", vbTextCompare)
    sFrPrd = Left(medPrdFr, iPos - 1) & Right(medPrdFr, Len(medPrdFr) - iPos)
    sToPrd = Left(medPrdTo, iPos - 1) & Right(medPrdTo, Len(medPrdTo) - iPos)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_RPTSOP030A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, FRPRICE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, gsUserID)
                Call SetSPPara(adcmdSave, 2, Change_SQLDate(wsDteTim))
                Call SetSPPara(adcmdSave, 3, Set_Quote(txtTitle.Text))
                Call SetSPPara(adcmdSave, 4, Set_Quote(medPrdFr))
                Call SetSPPara(adcmdSave, 5, Set_Quote(medPrdTo))
                Call SetSPPara(adcmdSave, 6, Set_Quote(sFrPrd))
                Call SetSPPara(adcmdSave, 7, Set_Quote(sToPrd))
                Call SetSPPara(adcmdSave, 8, To_Value(waResult(wiCtr, FRPRICE)))
                Call SetSPPara(adcmdSave, 9, To_Value(waResult(wiCtr, TOPRICE)))
                Call SetSPPara(adcmdSave, 10, IIf(wiCtr = waResult.UpperBound(1), "Y", "N"))
                Call SetSPPara(adcmdSave, 11, gsLangID)
                adcmdSave.Execute
            End If
        Next
    End If
    cnCon.CommitTrans
    
    'Create Stored Procedure String
    wsSQL = "EXEC usp_RPTSOP030 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "'"
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTSOP030"
    Else
    wsRptName = "RPTSOP030"
    End If
    
    NewfrmPrint.ReportID = "SOP030"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SOP030"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9
            Call cmdOK
            
        Case vbKeyF11
            Call cmdCancel
        
        Case vbKeyF12
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    MousePointer = vbHourglass
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Scr
    Call Ini_Grid

    MousePointer = vbDefault
End Sub

Private Sub Ini_Form()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "SOP030"
    
End Sub

Private Sub Ini_Scr()
    Me.Caption = wsFormCaption
    
    waResult.ReDim 0, -1, FRPRICE, TOPRICE
    Set tblDetail.Array = waResult
    'tblDetail.ReBind
    tblDetail.Bookmark = 0
    wiAction = DefaultPage

    tblCommon.Visible = False
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    
    Dim iCounter As Integer
    
    With waResult
        .ReDim 0, -1, FRPRICE, TOPRICE
        For iCounter = 0 To 9
            .AppendRows
            waResult(.UpperBound(1), FRPRICE) = iCounter * 10 + 1
            waResult(.UpperBound(1), TOPRICE) = (iCounter + 1) * 10
        Next
    End With
    
    tblDetail.Enabled = True
    tblDetail.ReBind
    'tblDetail.FirstRow = 0
    
    wgsTitle = "Sales Quantity/Amount History Report"
End Sub

Private Function InputValidation() As Boolean
    InputValidation = False
    
    If chk_medPrdFr = False Then
        Me.medPrdFr.SetFocus
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        Me.medPrdTo.SetFocus
        Exit Function
    End If
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, FRPRICE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = FRPRICE
                    tblDetail.SetFocus
                    Exit Function
                End If
                
                If Chk_NoDup2(wlCtr, waResult(wlCtr, FRPRICE), waResult(wlCtr, TOPRICE)) = False Then
                    tblDetail.Row = wlCtr - 1
                    tblDetail.Col = FRPRICE
                    tblDetail.SetFocus
                    Exit Function
                End If
                
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "定價範圍沒有來源資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tblDetail.Col = FRPRICE
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    InputValidation = True
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 5250
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set wcCombo = Nothing
   Set frmSOP030 = Nothing

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

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    
    With tblDetail
        .Columns(FRPRICE).Caption = Get_Caption(waScrItm, "FRPRICE")
        .Columns(TOPRICE).Caption = Get_Caption(waScrItm, "TOPRICE")
    End With
    
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblPrice.Caption = Get_Caption(waScrItm, "PRICE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case tcGo
            Call cmdOK
        Case tcCancel
            Call cmdCancel
        Case tcExit
            Unload Me
    End Select
    
End Sub

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 60, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        medPrdFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        
        tblDetail.SetFocus
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub

Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        gsMsg = "To must > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr.Text) = "/" Then
       gsMsg = "Must Input Period!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
       gsMsg = "Invalid Period!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
    
    chk_medPrdFr = True
End Function

Private Sub Ini_Grid()
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
     '   .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = FRPRICE To GDummy
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case FRPRICE
                    .Columns(wiCtr).Width = 1710
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = False
                Case TOPRICE
                    .Columns(wiCtr).Width = 1710
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = False
                Case GDummy
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
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

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case FRPRICE, TOPRICE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If ColIndex = FRPRICE Then
                    If Chk_grdFrPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                ElseIf ColIndex = TOPRICE Then
                    If Chk_grdToPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                End If
            End Select
            
            If .Columns(ColIndex).Text <> OldValue Then
                wbUpdate = True
            End If
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

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
        Case vbKeyF4        ' CALL COMBO BOX
            KeyCode = vbDefault
            'Call tblDetail_ButtonClick(.Col)
        
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
                Case TOPRICE
                    KeyCode = vbKeyDown
                    .Col = FRPRICE
                Case FRPRICE
                    KeyCode = vbDefault
                       .Col = TOPRICE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            Select Case .Col
                Case TOPRICE
                    .Col = FRPRICE
            End Select
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case FRPRICE
                       .Col = TOPRICE
            End Select
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    Select Case tblDetail.Col
        Case FRPRICE, TOPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = FRPRICE
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case FRPRICE
                    Call Chk_grdFrPrice(.Columns(FRPRICE).Text)
                    
                Case TOPRICE
                    Call Chk_grdToPrice(.Columns(TOPRICE).Text)
                'Case DisPer
                '    Call Chk_grdDisPer(.Columns(DisPer).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

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
        
        If Chk_grdFrPrice(waResult(LastRow, FRPRICE)) = False Then
                .Col = FRPRICE
                Exit Function
        End If
        
        If Chk_grdToPrice(waResult(LastRow, TOPRICE), waResult(LastRow, FRPRICE)) = False Then
                .Col = TOPRICE
                Exit Function
        End If
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Function Chk_NoDup2(ByRef inRow As Long, ByVal wsCurRec As String, ByVal wsCurRec1 As String) As Boolean
    Dim wlCtr As Long
     
    Chk_NoDup2 = False
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
            If (wsCurRec = waResult(wlCtr, FRPRICE)) And (wsCurRec1 = waResult(wlCtr, TOPRICE)) Then
                inRow = wlCtr
                gsMsg = "重覆項目!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                Exit Function
            End If
        End If
    Next
    
    Chk_NoDup2 = True
End Function

Private Function IsEmptyRow(Optional inRow) As Boolean
    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(FRPRICE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, FRPRICE)) = "" And _
                   Trim(waResult(inRow, TOPRICE)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
End Function

Private Function Chk_NoDup(inRow As Long) As Boolean
    Dim wlCtr As Long
    Dim wsCurRec As String
    
    Chk_NoDup = False
    
    wsCurRec = Format(tblDetail.Columns(FRPRICE), gsAmtFmt)
   
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           If wsCurRec = Format(waResult(wlCtr, FRPRICE), gsAmtFmt) Then
              gsMsg = "'由' 之定價與 '由' 之定價有重覆!"
              MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
              Exit Function
           End If
        End If
    Next
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
            If wsCurRec = Format(waResult(wlCtr, TOPRICE), gsAmtFmt) Then
                gsMsg = "'由' 之定價與 '至到' 之定價有重覆!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                Exit Function
            End If
        End If
    Next
    
    wsCurRec = Format(tblDetail.Columns(TOPRICE), gsAmtFmt)
   
    'For wlCtr = 0 To waResult.UpperBound(1)
    '    If inRow <> wlCtr Then
    '        If wsCurRec = Format(waResult(wlCtr, FRPRICE), gsAmtFmt) Then
    '            gsMsg = "'由' 之定價與 '至到' 之定價有重覆!"
    '            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    '            Exit Function
    '        End If
    '    End If
    'Next
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
            If wsCurRec = Format(waResult(wlCtr, TOPRICE), gsAmtFmt) Then
                gsMsg = "'至到' 之定價與 '至到' 之定價有重覆!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                Exit Function
            End If
        End If
    Next
    
    Chk_NoDup = True
End Function

Private Function Chk_grdFrPrice(inCode As String) As Boolean
    Chk_grdFrPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入範圍 '由' 定價之值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdFrPrice = False
        tblDetail.Col = FRPRICE
        Exit Function
    End If
End Function

Private Function Chk_grdToPrice(inCode As String, Optional inFrPrice As String) As Boolean
    Chk_grdToPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入範圍 '至' 定價之值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdToPrice = False
        Exit Function
    End If
    
    If IsMissing(inFrPrice) Then
        Exit Function
    End If
    
    If To_Value(inFrPrice) > To_Value(inCode) Then
        gsMsg = "範圍 '至' 定價之值大於 '由' 定價之值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdToPrice = False
        tblDetail.Col = TOPRICE
        Exit Function
    End If
End Function
