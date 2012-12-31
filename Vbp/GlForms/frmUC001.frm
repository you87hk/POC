VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmUC001 
   Caption         =   "UC001"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   11880
   Icon            =   "frmUC001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   12120
      OleObjectBlob   =   "frmUC001.frx":030A
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItmCodeTo 
      Height          =   300
      Left            =   6030
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   3000
   End
   Begin VB.ComboBox cboItmCodeFr 
      Height          =   300
      Left            =   2160
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   3000
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   7920
      Top             =   -120
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
            Picture         =   "frmUC001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUC001.frx":66AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   17160
      _ExtentX        =   30268
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
            Object.Visible         =   0   'False
            Key             =   "Open"
            Object.ToolTipText     =   "Open (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Key             =   "Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   7455
      Left            =   120
      OleObjectBlob   =   "frmUC001.frx":69C9
      TabIndex        =   2
      Top             =   1080
      Width           =   11775
   End
   Begin VB.Label lblItmCodeTo 
      Caption         =   "ITMCODETO"
      Height          =   225
      Left            =   5670
      TabIndex        =   5
      Top             =   645
      Width           =   375
   End
   Begin VB.Label lblItmCodeFr 
      Caption         =   "ITMCODEFR"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   650
      Width           =   1890
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
Attribute VB_Name = "frmUC001"
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

Private wsOldCusNo As String

Private wgsTitle As String

Private Const ITMCODE = 0
Private Const ITMDESC = 1
Private Const ITMWHSCODE = 2
Private Const ITMAVECOST = 3
Private Const ITMID = 4

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
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "MstWhsItem"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, ITMCODE, ITMID
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

    wbReadOnly = False
    
    Call SetButtonStatus("Default")
    Call SetFieldStatus("Default")
    Call SetFieldStatus("AfrActEdit")
    
    'lblDspYear = Left(gsSystemDate, 4)
    
    wlKey = 0
    wlLineNo = 1
    wsTrnCd = ""
        
    'tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    'FocusMe cboJobCode
    
    'Ini_Scr_AfrKey
    'tblDetail.Col = ITMAVECOST
    'tblDetail.ScrollBars = dbgVertical
    'cboItmCodeFr.SetFocus
    FocusMe cboItmCodeFr
End Sub

Private Sub Ini_Scr_AfrKey()
    If LoadRecord() = False Then
        wiAction = AddRec
        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        'If RowLock(wsConnTime, wsKeyType, cboJobCode.Text, wsFormID, wsUsrId) = False Then
        '    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
        '    MsgBox gsMsg, vbOKOnly, gsTitle
        '    tblDetail.ReBind
        'End If

         Call SetButtonStatus("AfrKeyEdit")
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    Call SetFieldStatus("AfrKey")
    
    'If tblDetail.Enabled = True Then
    '    tblDetail.Col = ITMWHSCODE
    '    tblDetail.SetFocus
    'End If
    
  '      wiAction = AddRec
  '      Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
  '      Call SetButtonStatus("AfrKeyAdd")
  '      Call SetFieldStatus("AfrKey")
        
  '      cboSaleCode.SetFocus
End Sub

Private Sub cboItmCodeTo_LostFocus()
    FocusMe cboItmCodeTo, True
End Sub

Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        'Set wcCombo = cboCusCode
        'Call cboCusCode_DropDown
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageDown
            'KeyCode = 0
            'If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
            '    tabDetailInfo.Tab = tabDetailInfo.Tab + 1
            '    Exit Sub
            'End If
            
        Case vbKeyPageUp
            'KeyCode = 0
            'If tabDetailInfo.Tab > 0 Then
            '    tabDetailInfo.Tab = tabDetailInfo.Tab - 1
            '    Exit Sub
            'End If
       
        Case vbKeyF6
            Call cmdOpen
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
        
        'Case vbKeyF3
        '    If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF7
            If tbrProcess.Buttons(tcRefresh).Enabled = True Then Call cmdRefresh
            
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
    lblItmCodeFr.Caption = Get_Caption(waScrItm, "ITMCODEFR")
    lblItmCodeTo.Caption = Get_Caption(waScrItm, "ITMCODETO")
    
    With tblDetail
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "ITMCODE")
        .Columns(ITMDESC).Caption = Get_Caption(waScrItm, "ITMDESC")
        .Columns(ITMWHSCODE).Caption = Get_Caption(waScrItm, "ITMWHSCODE")
        .Columns(ITMAVECOST).Caption = Get_Caption(waScrItm, "ITMAVECOST")
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
   
    wsActNam(1) = Get_Caption(waScrItm, "UCADD")
    wsActNam(2) = Get_Caption(waScrItm, "UCEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "UCDELETE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
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
    Set frmUC001 = Nothing
End Sub

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

    'If wiAction <> AddRec Then
    '    If ReadOnlyMode(wsConnTime, wsKeyType, cboJobCode.Text, wsFormID) Or wbReadOnly Then
    '        gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
    '        MsgBox gsMsg, vbOKOnly, gsTitle
    '        MousePointer = vbDefault
    '        Exit Function
    '    End If
    'End If
   
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Function
    End If
    
    '' Last Check when Add
    'If wiAction = AddRec Then
    '    If Chk_KeyExist() = True Then
    '        Call GetNewKey
    '    End If
    'End If
         
   ' If lblDspNetAmtOrg.Caption > Get_CreditLimit(wlCusID, wlKey, Trim(medDocDate.Text)) Then
   '    gsMsg = "已超過信貸額!"
   '    MsgBox gsMsg, vbOKOnly, gsTitle
   '    MousePointer = vbDefault
   '    Exit Function
   ' End If
    
    wlRowCtr = waResult.UpperBound(1)
    'wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_UC001"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, waResult(wiCtr, ITMID))
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, ITMWHSCODE))
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, ITMAVECOST))
                Call SetSPPara(adcmdSave, 4, gsUserID)
                Call SetSPPara(adcmdSave, 5, wsGenDte)
                adcmdSave.Execute
            End If
        Next
    End If
    cnCon.CommitTrans
    
    'If wiAction = AddRec Then
    '    If Trim(wsDocNo) <> "" Then
    '        Call cmdPrint(wsDocNo)
    '        gsMsg = "文件號 : " & wsDocNo & " 已製成!"
    '        MsgBox gsMsg, vbOKOnly, gsTitle
    '    Else
    '        gsMsg = "文件儲存失敗!"
    '        MsgBox gsMsg, vbOKOnly, gsTitle
    '    End If
    'End If
    
    If wiAction = CorRec Then
        gsMsg = "文件已儲存!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Call cmdCancel
    Call SetButtonStatus("AfrKeyEdit")
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
    cmdDel = True
End Function

Private Function InputValidation() As Boolean
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, ITMAVECOST)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = ITMAVECOST
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
            'For wlCtr1 = 0 To .UpperBound(1)
            '    If wlCtr <> wlCtr1 Then
            '        If waResult(wlCtr, BOOKCODE) = waResult(wlCtr1, BOOKCODE) Then
            '          gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
            '          MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            '          tblDetail.Col = BOOKCODE
            '          tblDetail.SetFocus
            '          Exit Function
            '        End If
            '    End If
            'Next
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tblDetail.Col = ITMAVECOST
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    'If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
    '    tblDetail.FirstRow = tblDetail.Row
    '    tblDetail.Col = BOOKCODE
    '    tblDetail.SetFocus
    '    Exit Function
    'End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cmdNew()
    Dim newForm As New frmUC001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show
End Sub

Private Sub cmdOpen()
    Dim newForm As New frmUC001
    
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
    wsFormID = "UC001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    
    Call LoadWSINFO
End Sub

Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("Default")
    
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

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = ITMCODE To ITMID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case ITMCODE
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case ITMDESC
                    .Columns(wiCtr).Width = 6000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case ITMWHSCODE
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case ITMAVECOST
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case ITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
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

    'If ColIndex = COSTCODE Then
    '    Call LoadJobCost(sTemp)
    'End If
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsCostCode As String
    Dim wsCostDesc As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    'If tblCommon.Visible = True Then
    '    Cancel = False
    '    tblDetail.Columns(ColIndex).Text = OldValue
    '    Exit Sub
    'End If
       
    With tblDetail
        Select Case ColIndex
            Case ITMAVECOST
                'If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                '    GoTo Tbl_BeforeColUpdate_Err
                'End If
                
                If Chk_grdBal(.Columns(ITMAVECOST).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If

                .Columns(ITMAVECOST).Text = Format(.Columns(ITMAVECOST).Text, gsAmtFmt)
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
                Case ITMAVECOST
                    KeyCode = vbKeyDown
                    .Col = ITMAVECOST
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
        Case vbKeyRight
            KeyCode = vbDefault
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    Select Case tblDetail.Col
        Case ITMAVECOST
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        'Case Price, DisPer
        '    Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
    End Select
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = ITMAVECOST
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case ITMAVECOST
                    Call Chk_grdBal(.Columns(ITMAVECOST).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdBal(inBal As String) As Boolean
    Chk_grdBal = False
    
    If Trim(inBal) = "" Then
        gsMsg = "必需輸入數值!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBal = False
        Exit Function
    End If

    'If To_Value(inBal) < 0 Then
    '    gsMsg = "數量必需大過或等於零!"
    '    MsgBox gsMsg, vbOKOnly, gsTitle
    '    Chk_grdBal = False
    '    Exit Function
    'End If

    Chk_grdBal = True
End Function

Private Function IsEmptyRow(Optional inRow) As Boolean
    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(ITMAVECOST)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, ITMCODE)) = "" And _
                   Trim(waResult(inRow, ITMAVECOST)) = "" Then
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
        
        If Chk_grdBal(waResult(LastRow, ITMAVECOST)) = False Then
            .Col = ITMAVECOST
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
        
        gsMsg = "你是否確定要儲存現時之作業?"
        If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
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

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            cboItmCodeFr.Enabled = True
            cboItmCodeTo.Enabled = True
            Me.tblDetail.Enabled = True
            
        Case "AfrActAdd"
            Me.tblDetail.Enabled = True
       
        Case "AfrActEdit"
            Me.tblDetail.Enabled = True
        
        Case "AfrKey"
            cboItmCodeFr.Enabled = True
            cboItmCodeTo.Enabled = True
            Me.tblDetail.Enabled = True
    End Select
End Sub

Private Sub LoadWSINFO()
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT * FROM sysWSINFO WHERE WSID ='" + gsWorkStationID + "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        If gsLangID = "2" Then
            wgsTitle = ReadRs(rsRcd, "WSCTITLE")
        Else
            wgsTitle = ReadRs(rsRcd, "WSTITLE")
        End If
    Else
        wgsTitle = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False

    If gsLangID = 1 Then
        wsSQL = "SELECT WhsItemID, ItmCode, ItmEngName ITMDESC, WhsItemWhsCode, WhsItemAveCost "
        wsSQL = wsSQL & "FROM  MstItem, MstWhsItem "
        wsSQL = wsSQL & "WHERE WhsItemItmID = ItmID "
        wsSQL = wsSQL & "AND ItmCode >= '" & cboItmCodeFr & "'"
        wsSQL = wsSQL & "AND ItmCode <= '" & cboItmCodeTo & "'"
        wsSQL = wsSQL & "AND ItmID = WhsItemID "
        wsSQL = wsSQL & "AND WhsItemStatus = '1' "
        wsSQL = wsSQL & "ORDER BY ItmCode"
    Else
        wsSQL = "SELECT WhsItemID, ItmCode, ItmChiName ITMDESC, WhsItemWhsCode, WhsItemAveCost "
        wsSQL = wsSQL & "FROM  MstItem, MstWhsItem "
        wsSQL = wsSQL & "WHERE WhsItemItmID = ItmID "
        wsSQL = wsSQL & "AND ItmCode >= '" & cboItmCodeFr & "'"
        wsSQL = wsSQL & "AND ItmCode <= '" & cboItmCodeTo & "'"
        wsSQL = wsSQL & "AND ItmID = WhsItemID "
        wsSQL = wsSQL & "AND WhsItemStatus = '1' "
        wsSQL = wsSQL & "ORDER BY ItmCode"
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, ITMCODE, ITMID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), ITMDESC) = ReadRs(rsRcd, "ITMDESC")
             waResult(.UpperBound(1), ITMWHSCODE) = ReadRs(rsRcd, "WhsItemWhsCode")
             waResult(.UpperBound(1), ITMAVECOST) = ReadRs(rsRcd, "WhsItemAveCost")
             waResult(.UpperBound(1), ITMID) = ReadRs(rsRcd, "WhsItemID")
             rsRcd.MoveNext
         Loop
         'wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    LoadRecord = True
    
End Function

Private Sub mnuPopUpSub_Click(Index As Integer)
    'Call Call_PopUpMenu(waPopUpSub, Index)
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
End Sub

Private Function cmdRefresh() As Boolean
    cmdRefresh = False
    cmdRefresh = True
End Function

Private Sub cboItmCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeFr
    
    If gsLangID = "1" Then
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmEngName, ItmClass FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    Else
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmChiName, ItmClass FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    End If
    Call Ini_Combo(4, wsSQL, cboItmCodeFr.Left, cboItmCodeFr.Top + cboItmCodeFr.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeFr_GotFocus()
    FocusMe cboItmCodeFr
    Set wcCombo = cboItmCodeFr
End Sub

Private Sub cboItmCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeFr, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmCodeFr.Text) <> "" And _
            Trim(cboItmCodeTo.Text) = "" Then
            cboItmCodeTo.Text = cboItmCodeFr.Text
        End If
        
        cboItmCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmCodeFr_LostFocus()
    FocusMe cboItmCodeFr, True
End Sub

Private Sub cboItmCodeTo_DropDown()
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeTo
  
    If gsLangID = "1" Then
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmEngName, ItmClass FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    Else
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmChiName, ItmClass FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    End If
    Call Ini_Combo(4, wsSQL, cboItmCodeTo.Left, cboItmCodeTo.Top + cboItmCodeTo.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeTo_GotFocus()
    FocusMe cboItmCodeTo
    Set wcCombo = cboItmCodeTo
End Sub

Private Sub cboItmCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeTo, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmCodeTo = False Then
            cboItmCodeTo.SetFocus
            Exit Sub
        End If
        
        Ini_Scr_AfrKey
        
        If tblDetail.Enabled = True Then
            tblDetail.SetFocus
        Else
            cboItmCodeFr.SetFocus
        End If
    End If
End Sub

Private Function chk_cboItmCodeTo() As Boolean
    chk_cboItmCodeTo = False
    
    If UCase(cboItmCodeFr.Text) > UCase(cboItmCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmCodeTo = True
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If

End Sub

