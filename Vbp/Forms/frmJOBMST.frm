VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmJOBMST 
   Caption         =   "書本對換價"
   ClientHeight    =   6015
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   6360
   Icon            =   "frmJOBMST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6360
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   4560
      OleObjectBlob   =   "frmJOBMST.frx":08CA
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtJobCost 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cboJobNo 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   1560
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
            Picture         =   "frmJOBMST.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOBMST.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmJOBMST.frx":6C6D
      TabIndex        =   2
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label lblJobCost 
      Caption         =   "PORTNO"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label lblJobNo 
      Caption         =   "ITMCODE"
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
      TabIndex        =   3
      Top             =   540
      Width           =   1455
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
Attribute VB_Name = "frmJOBMST"
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

Private Const GSONO = 0
Private Const GNETAMTL = 1
Private Const GDOCID = 2


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
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "MstJobNo"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    waResult.ReDim 0, -1, GSONO, GDOCID
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
    
    wlCusID = 0
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    
    FocusMe cboJobNo
    
    
End Sub


Private Sub Ini_Scr_AfrKey()

If LoadRecord = False Then Exit Sub

    
    wiAction = CorRec
    If RowLock(wsConnTime, wsKeyType, cboJobNo, wsFormID, wsUsrId) = False Then
        gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
        MsgBox gsMsg, vbOKOnly, gsTitle
        tblDetail.ReBind
       
    End If
    
    Call SetButtonStatus("AfrKeyEdit")
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    Call SetFieldStatus("AfrKey")
    
    txtJobCost.SetFocus
    
    'If tblDetail.Enabled = True Then
    '    tblDetail.SetFocus
    'End If
    
End Sub

Private Sub cboJobNo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboJobNo
  
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM SOASOHD, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboJobNo.SelLength > 0, "", Set_Quote(cboJobNo.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    Call Ini_Combo(3, wsSQL, cboJobNo.Left, cboJobNo.Top + cboJobNo.Height, tblCommon, "JOBMST", "TBLJOBNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboJobNo_GotFocus()
    FocusMe cboJobNo
End Sub

Private Sub cboJobNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboJobNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboJobNo() = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If
End Sub

Private Sub cboJobNo_LostFocus()
    FocusMe cboJobNo, True
End Sub

Private Function Chk_cboJobNo() As Boolean
    Dim wsStatus As String
    
    Chk_cboJobNo = False
    
    If Trim(cboJobNo.Text) = "" Then
        gsMsg = "必需輸入工程編號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboJobNo.SetFocus
        Exit Function
    End If
        
    If Chk_JobNo(cboJobNo) = False Then
        gsMsg = "工程編號不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboJobNo.SetFocus
        Exit Function
    End If
    
    Chk_cboJobNo = True
End Function


Public Function Chk_JobNo(ByVal inJobNo As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "Select SOHDDOCNO "
    wsSQL = wsSQL & " From SOASOHD "
    wsSQL = wsSQL & " Where SOHDDOCNO = '" & Set_Quote(inJobNo) & "'"
    wsSQL = wsSQL & " AND SOHDSTATUS IN ('1','4') "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_JobNo = True
    Else
        Chk_JobNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function




Private Sub mnuPopUpSub_Click(Index As Integer)
  Call Call_PopUpMenu(waPopUpSub, Index)
End Sub

Private Sub txtJobCost_GotFocus()

    FocusMe txtJobCost
    
End Sub

Private Sub txtJobCost_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtJobCost.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtJobCost Then
            
            If tblDetail.Enabled = True Then
                tblDetail.SetFocus
            End If
    
        End If
    End If

End Sub

Private Function chk_txtJobCost() As Boolean
    
    chk_txtJobCost = False
    
    
    If To_Value(txtJobCost.Text) <= 0 Then
        gsMsg = "錯誤!一定大於零"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtJobCost.SetFocus
        Exit Function
    End If
    
    txtJobCost.Text = Format(txtJobCost.Text, gsAmtFmt)
    
    chk_txtJobCost = True
    
End Function

Private Sub txtJobCost_LostFocus()
txtJobCost.Text = Format(txtJobCost.Text, gsAmtFmt)
FocusMe txtJobCost, True
End Sub



Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        Set wcCombo = cboJobNo
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
   Select Case KeyCode
        
        
        Case vbKeyF6
            Call cmdOpen
        
        
        'Case vbKeyF2
        '    If wiAction = DefaultPage Then Call cmdNew
            
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
       
        
        'Case vbKeyF3
        '    If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF9
        
        'If tbrProcess.Buttons(tcFind).Enabled = True Then
        '    Call cmdFind
        'End If
        
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
    Dim wiCtr As Long
    
    LoadRecord = False
    
    
    wsSQL = "SELECT SOHDDOCID, SOHDDOCNO, SOHDNETAMTL, JOBCOST, SOHDCUSID "
    wsSQL = wsSQL & "FROM SOASOHD, MSTJOBNO "
    wsSQL = wsSQL & "WHERE SOHDJOBNO = '" & Set_Quote(cboJobNo.Text) & "' "
    wsSQL = wsSQL & "AND SOHDJOBNO *= JOBNO "


    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        txtJobCost.Text = Format("0", gsAmtFmt)
        
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlCusID = ReadRs(rsRcd, "SOHDCUSID")
    txtJobCost = Format(To_Value(ReadRs(rsRcd, "JOBCOST")), gsAmtFmt)
        
        
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, GSONO, GDOCID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), GSONO) = ReadRs(rsRcd, "SOHDDOCNO")
             waResult(.UpperBound(1), GNETAMTL) = Format(To_Value(ReadRs(rsRcd, "SOHDNETAMTL")), gsAmtFmt)
             waResult(.UpperBound(1), GDOCID) = ReadRs(rsRcd, "SOHDDOCID")
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
        
        
    rsRcd.Close
    Set rsRcd = Nothing
  
 
 LoadRecord = True
 
End Function


Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblJobNo.Caption = Get_Caption(waScrItm, "JOBNO")
    lblJobCost.Caption = Get_Caption(waScrItm, "JOBCOST")
    
    With tblDetail
        .Columns(GSONO).Caption = Get_Caption(waScrItm, "GSONO")
        .Columns(GNETAMTL).Caption = Get_Caption(waScrItm, "GNETAMTL")
    End With
    
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "IPADD")
    wsActNam(2) = Get_Caption(waScrItm, "IPEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "IPDELETE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 6420
        Me.Width = 6480
    End If
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
    Set frmJOBMST = Nothing
End Sub


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

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlRowCtr As Long
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboJobNo, wsFormID) Then
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
    
    wlRowCtr = waResult.UpperBound(1)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
      
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_JOBMST"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GSONO)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, cboJobNo)
                Call SetSPPara(adcmdSave, 3, txtJobCost)
                Call SetSPPara(adcmdSave, 4, wiCtr)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, GDOCID))
                Call SetSPPara(adcmdSave, 6, gsUserID)
                Call SetSPPara(adcmdSave, 7, wsGenDte)
                
                adcmdSave.Execute
                
            End If
        Next
    End If
    cnCon.CommitTrans
    
    Set adcmdSave = Nothing
    cmdSave = True
    
    gsMsg = "工程已儲存!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
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
      
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    
    If Not chk_txtJobCost Then Exit Function
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, GSONO)) <> "" Then
                
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
                
                For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                If waResult(wlCtr, GSONO) = waResult(wlCtr1, GSONO) Then
                  gsMsg = "工程已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
                End If
                End If
                Next wlCtr1
                
                
                
            End If
        Next wlCtr
    End With
    
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有設定工程!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
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

    Dim newForm As New frmJOBMST
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()

    Dim newForm As New frmJOBMST
    
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
    wsFormID = "JOBMST"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "JB"
End Sub

Private Sub cmdCancel()
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboJobNo.SetFocus
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
        If Chk_GrdRow(To_Value(.Bookmark)) = False Then
            Cancel = True
            Exit Sub
        End If
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
        'Case tcDelete
        '    Call cmdDel
        Case tcSave
            Call cmdSave
        Case tcCancel
           If tbrProcess.Buttons(tcSave).Enabled = True Then
           If MsgBox("Are you sure to cancel this operation?", vbYesNo, gsTitle) = vbYes Then
                Call cmdCancel
           End If
           Else
                Call cmdCancel
           End If
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
        
        For wiCtr = GSONO To GDOCID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case GSONO
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 15
                    
                Case GNETAMTL
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                    
                Case GDOCID
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
    Dim wdNetAmtl As Double
    Dim wsSoId As String
    
    On Error GoTo tblDetail_BeforeColUpdate_Err
    

    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
    
    With tblDetail
        Select Case ColIndex
            Case GSONO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdSoNo(.Columns(ColIndex).Text, wsSoId, wdNetAmtl) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(GDOCID).Text = wsSoId
                .Columns(GNETAMTL).Text = Format(wdNetAmtl, gsAmtFmt)
                
                
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
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err

    With tblDetail
        Select Case ColIndex
            Case GSONO
                wsSQL = "SELECT SOHDDOCNO FROM SOASOHD "
                wsSQL = wsSQL & " WHERE SOHDSTATUS <> '2' "
                wsSQL = wsSQL & " AND SOHDCUSID = " & wlCusID & " "
                wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(GSONO).Text) & "%' "
                If waResult.UpperBound(1) > -1 Then
                          wsSQL = wsSQL & " AND SOHDDOCNO NOT IN ( "
                          For wiCtr = 0 To waResult.UpperBound(1)
                                wsSQL = wsSQL & " '" & Set_Quote(waResult(wiCtr, GSONO)) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                          Next
                End If
                
                wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
                
                Call Ini_Combo(1, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLJOBNO", Me.Width, Me.Height)
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
            If .Bookmark = waResult.UpperBound(1) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case vbKeyF8        ' DELETE LINE
            KeyCode = vbDefault
            If IsNull(.Bookmark) Then Exit Sub
            If .EditActive = True Then Exit Sub
            gsMsg = "你是否確認要刪除此列?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then Exit Sub
            .Delete
            .Update
            If .Row = -1 Then
                .Row = 0
            End If
            .Refresh
            .SetFocus

        Case vbKeyReturn
            If .Col <> GNETAMTL Then
            KeyCode = vbDefault
                  .Col = .Col + 1
            Else
            KeyCode = vbKeyDown
            .Col = GSONO
            
            End If
            
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> GSONO Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> GNETAMTL Then
                  .Col = .Col + 1
            End If
            
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
           .Col = GSONO
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case GSONO
                    Call Chk_grdSoNo(.Columns(GSONO).Text, "", 0)
                    
                                   
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdSoNo(inSoNo As String, OutSoID As String, outNetAmtl As Double) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset


    wsSQL = "SELECT SOHDDOCID, SOHDNETAMTL FROM SOASOHD "
    wsSQL = wsSQL & " WHERE SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSQL = wsSQL & " And SOHDSTATUS <> '2' "
    wsSQL = wsSQL & " And SOHDCUSID = " & wlCusID
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        OutSoID = ReadRs(rsDes, "SOHDDOCID")
        outNetAmtl = ReadRs(rsDes, "SOHDNETAMTL")
          
        Chk_grdSoNo = True
    Else
        OutSoID = ""
        outNetAmtl = 0
        gsMsg = "沒有此工程!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdSoNo = False
        
    End If
    rsDes.Close
    Set rsDes = Nothing
    
    
End Function

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(GSONO)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, GSONO)) = "" And _
                   Trim(waResult(inRow, GNETAMTL)) = "" And _
                   Trim(waResult(inRow, GDOCID)) = "" Then
                   Exit Function
                End If
            End If
        End If
        
    
    IsEmptyRow = False
    
End Function
Private Function Chk_GrdRow(ByVal LastRow As Long) As Boolean
    Dim wlCtr As Long
    
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
        
        If Chk_grdSoNo(waResult(LastRow, GSONO), "", 0) = False Then
            .Col = GSONO
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
            'If wiAction = DelRec Then
            '    If cmdDel = True Then
            '        Exit Function
            '    End If
            'Else
                If cmdSave = True Then
                    Exit Function
                End If
            'End If
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
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "ReadOnly"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcExit).Enabled = True
            
            End With
    End Select
End Sub

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
        
            Me.cboJobNo.Enabled = False
            Me.txtJobCost.Enabled = False
            Me.tblDetail.Enabled = False
            
            
        Case "AfrActAdd"
        
            Me.cboJobNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboJobNo.Enabled = True
        
        Case "AfrKey"
            Me.cboJobNo.Enabled = False
            Me.txtJobCost.Enabled = True
            
            If wiAction = CorRec Then
                Me.tblDetail.Enabled = True
                
            End If
    End Select
End Sub


Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(GSONO)
    
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, GSONO) Then
                  gsMsg = "工程已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function


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


