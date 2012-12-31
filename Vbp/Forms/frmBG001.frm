VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmBG001 
   Caption         =   "Book List Maintenance"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   9795
   Icon            =   "frmBG001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11040
      OleObjectBlob   =   "frmBG001.frx":030A
      TabIndex        =   2
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
      Height          =   5655
      Left            =   240
      OleObjectBlob   =   "frmBG001.frx":2A0D
      TabIndex        =   1
      Top             =   1440
      Width           =   11055
   End
   Begin VB.Frame fra1 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtRemark 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   11
         Top             =   600
         Width           =   9255
      End
      Begin VB.Label lblDspItmLstUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   8520
         TabIndex        =   10
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblDspItmLstUpdUsr 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   9
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label lblItmLstUpdDate 
         Caption         =   "ITMLSTUPDDATE"
         Height          =   240
         Left            =   6960
         TabIndex        =   8
         Top             =   6885
         Width           =   1500
      End
      Begin VB.Label lblItmLstUpdUsr 
         Caption         =   "ITMLSTUPDUSR"
         Height          =   240
         Left            =   360
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
            Picture         =   "frmBG001.frx":7CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":85A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":8E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":92D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":9724
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":9A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":9E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":A2E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":A5FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":A916
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":AD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBG001.frx":B644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmBG001"
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
Private wsDueDate As String
Private wsPayCode As String
Private wsWhsCode As String
Private wsMethodCode As String
Private wsCurCode As String
Private wdExcr As Double
Private wgsTitle As String
Private wsShipName As String
Private wsShipAdr1 As String
Private wsShipAdr2 As String
Private wsShipAdr3 As String
Private wsShipAdr4 As String

Private Const BOOKCODE = 0
Private Const BARCODE = 1
Private Const BOOKNAME = 2
Private Const BOOKID = 3

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
Private wiRevNo As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlCusTyp As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "MstItemList"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, BOOKCODE, BOOKID
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
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    FocusMe cboDocNo
End Sub

Private Sub Ini_Scr_AfrKey()
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
    
    
    Call SetFieldStatus("AfrKey")
    
    txtRemark.SetFocus
    
  '      wiAction = AddRec
  '      Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
  '      Call SetButtonStatus("AfrKeyAdd")
  '      Call SetFieldStatus("AfrKey")
        
  '      cboSaleCode.SetFocus
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
    
    With tblDetail
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
    End With
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
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
    
    'If chk_cboCusCode = False Then
    '    Exit Function
    'End If
    
    'If Chk_medDocDate = False Then
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
        
    adcmdSave.CommandText = "USP_BL001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, txtRemark)
    Call SetSPPara(adcmdSave, 6, wsFormID)
    Call SetSPPara(adcmdSave, 7, gsWorkStationID)
    Call SetSPPara(adcmdSave, 8, gsUserID)
    Call SetSPPara(adcmdSave, 9, wsGenDte)
      
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 10)
    wsDocNo = GetSPPara(adcmdSave, 11)
    
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
        
    adcmdDelete.CommandText = "USP_BL001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, "")          'txtRemark
    Call SetSPPara(adcmdDelete, 6, wsFormID)
    Call SetSPPara(adcmdDelete, 7, gsWorkStationID)
    Call SetSPPara(adcmdDelete, 8, gsUserID)
    Call SetSPPara(adcmdDelete, 9, wsGenDte)
    
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 10)
    wsDocNo = GetSPPara(adcmdDelete, 11)
    
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
    
    'If Not chk_txtRevNo Then Exit Function
    'If Not Chk_medDocDate Then Exit Function
    'If Not chk_cboCusCode() Then Exit Function
    'If Not Chk_cboSaleCode Then Exit Function
    'If Not Chk_cboMethodCode Then Exit Function
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, BOOKCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
            For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                    If waResult(wlCtr, BOOKCODE) = waResult(wlCtr1, BOOKCODE) Then
                      gsMsg = "重覆書本!"
                      MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                      tblDetail.SetFocus
                      Exit Function
                    End If
                End If
            Next
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "訂購單沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
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
    
    Call LoadWSINFO
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

Private Sub Get_DefVal()
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSql As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstCUSTOMER "
    wsSql = wsSql & "WHERE CUSID = " & wlCusID
    rsDefVal.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        wlSaleID = ReadRs(rsDefVal, "CUSSALEID")
        wlCusTyp = To_Value(ReadRs(rsDefVal, "CUSTYPID"))
        wsPayCode = ReadRs(rsDefVal, "CUSPAYCODE")
        lblDspCusContact = ReadRs(rsDefVal, "CUSCONTACTPERSON")
        lblDspCusAddress = ReadRs(rsDefVal, "CUSADDRESS1") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS2") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS3") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS4")
        wsShipName = ReadRs(rsDefVal, "CUSNAME")
        wsShipAdr1 = ReadRs(rsDefVal, "CUSADDRESS1")
        wsShipAdr2 = ReadRs(rsDefVal, "CUSADDRESS2")
        wsShipAdr3 = ReadRs(rsDefVal, "CUSADDRESS3")
        wsShipAdr4 = ReadRs(rsDefVal, "CUSADDRESS4")
        
          Else
        wlSaleID = 0
        wlCusTyp = 0
        wsPayCode = ""
        wsShipName = ""
        wsShipAdr1 = ""
        wsShipAdr2 = ""
        wsShipAdr3 = ""
        wsShipAdr4 = ""
        
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
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = BOOKCODE To BOOKID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case BOOKCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                Case BARCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case BOOKNAME
                    .Columns(wiCtr).Width = 4000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = False
                Case BOOKID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
        .Styles("EvenRow").BackColor = &H8000000F
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
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double

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
                
                If Chk_grdBookCode(.Columns(ColIndex).Text, wsBookID, wsBookCode, wsBarCode, wsBookName) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(BOOKID).Text = wsBookID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(BOOKNAME).Text = wsBookName
                
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
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM FROM mstITEM "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    
                    If waResult.UpperBound(1) > -1 Then
                        wsSql = wsSql & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                        Next
                    End If
                    
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM FROM mstITEM "
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
                       .Col = .Col + 1
                Case BARCODE
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case BOOKNAME
                    KeyCode = vbDefault
                    .Col = BOOKCODE
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case BARCODE, BOOKNAME
                    .Col = .Col - 1
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case BOOKCODE, BARCODE
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
           .Col = BOOKCODE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(BOOKCODE).Text, "", "", "", "")
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdBookCode(inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    wsSql = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMBARCODE FROM MSTITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        OutBarCode = ReadRs(rsDes, "ITMBARCODE")
        
        Chk_grdBookCode = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
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
        
        If Chk_grdBookCode(waResult(LastRow, BOOKCODE), "", "", "", "") = False Then
            .Col = BOOKCODE
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
            Me.txtRemark.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
           Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
           Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
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
                  gsMsg = "重覆書本!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function

Private Function Chk_NoDup2(inItmCode As String, inWhsCode As String) As Boolean
' CHECK NEW ENTRY FRO DUPLICATES
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    
    Chk_NoDup2 = False
    
    If waResult.UpperBound(1) = -1 Then
       Chk_NoDup2 = True
       Exit Function
    End If
    
    If Trim(inItmCode) = "" Then Exit Function
    
   ' If optStlMtd(0).Value = True Then
   '     For wlCtr = 0 To waInvoice.UpperBound(1)
   '
   '         If inInvNo = waInvoice(wlCtr, Tab1InvNo) And _
   '            inInvLn = waInvoice(wlCtr, Tab1InvLn) Then
   '            Call Dsp_Err("E0014", "", "E", Me.Caption)
   '           Exit Function
   '        End If
   '    Next
   ' Else
        For wlCtr = 0 To waResult.UpperBound(1)
            If inItmCode = waResult(wlCtr, BOOKCODE) Then
                gsMsg = "重覆書本!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
               Exit Function
            End If
        Next
    
    'End If
    Chk_NoDup2 = True

End Function

Private Sub cmdPrint(InDocNo As String)
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
    wsSql = "EXEC usp_RPTSN001 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & InDocNo & "', "
    wsSql = wsSql & "'" & InDocNo & "', "
    wsSql = wsSql & "'" & "" & "', "
    wsSql = wsSql & "'" & String(10, "z") & "', "
    wsSql = wsSql & "'" & String(6, "0") & "', "
    wsSql = wsSql & "'" & String(6, "9") & "', "
    wsSql = wsSql & "'" & "N" & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then wsRptName = "C" + "RPTSN001"
    
    NewfrmPrint.ReportID = "SN001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SN001"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub LoadWSINFO()
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT * FROM sysWSINFO WHERE WSID ='" + gsWorkStationID + "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
    
    
    wsWhsCode = ReadRs(rsRcd, "WSWHSCODE")
    wsMethodCode = ReadRs(rsRcd, "WSMETHODCODE")
    wsCurCode = ReadRs(rsRcd, "WSCURR")
    wdExcr = To_Value(ReadRs(rsRcd, "WSEXCR"))
    If gsLangID = "2" Then
    wgsTitle = ReadRs(rsRcd, "WSCTITLE")
    Else
    wgsTitle = ReadRs(rsRcd, "WSTITLE")
    End If
    
    
    Else
    
    wsWhsCode = ""
    wsMethodCode = ""
    wsCurCode = wsBaseCurCd
    wdExcr = 1
    wgsTitle = ""
    
    
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    wsSql = "SELECT ITMLSTDOCID, ITMLSTDOCNO, ITMLSTRMK, ITMLSTDTITEMID, ITMCODE, ITMBARCODE, ITMCHINAME "
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
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, BOOKCODE, BOOKID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITMCHINAME")
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "ITMLSTDTITEMID")
             rsRcd.MoveNext
         Loop
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

    
    wsSql = "SELECT SNHDSTATUS FROM soaSNHD WHERE SNHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
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
        .TableKey = "SnHdDocNo"
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
    Dim wsBarCode As String
    Dim wsSeries As String
    
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
       wsBarCode = ReadRs(rsRcd, "ITMBARCODE")
       
       With waResult
             .AppendRows
             waResult(.UpperBound(1), BOOKCODE) = wsISBN
             waResult(.UpperBound(1), BARCODE) = wsBarCode
             waResult(.UpperBound(1), BOOKNAME) = wsName
             waResult(.UpperBound(1), BOOKID) = wsItmID
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

Public Function Chk_ItmLstDocNo(ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT ITMLSTSTATUS FROM MSTITEMLIST WHERE  ITMLSTDOCNO= '" & Set_Quote(InDocNo) & "'"
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

Private Sub txtRemark_GotFocus()
    FocusMe txtRemark
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtRemark, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tblDetail.SetFocus
    End If
End Sub

Private Sub txtRemark_LostFocus()
    FocusMe txtRemark, True
End Sub

