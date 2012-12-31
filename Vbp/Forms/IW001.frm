VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmIW001 
   Caption         =   "書本對換價"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "IW001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboWhsCode 
      Height          =   300
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cboITMCODE 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "IW001.frx":08CA
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   1320
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
            Picture         =   "IW001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IW001.frx":6945
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
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "IW001.frx":6C6D
      TabIndex        =   6
      Top             =   1440
      Width           =   9495
   End
   Begin VB.Label lblWhsCode 
      Caption         =   "WHSCODE"
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
      Left            =   4680
      TabIndex        =   8
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblDspItmName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label lblItemChiName 
      Caption         =   "VDRITEMCHINAME"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblItmCode 
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
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PoP Up"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmIW001"
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


Private Const GITMCODE = 0
Private Const GWHSCODE = 1
Private Const GITEMNAME = 2
Private Const GLOTNO = 3
Private Const GLEAD = 4
Private Const GAVGCOST = 5
Private Const GPOCOST = 6
Private Const GLPRICE = 7
Private Const GITMID = 8
Private Const GDummy = 9


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
Private wlItmID As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "mstWhsItem"
Private wsFormID As String
Private wsUsrId As String



Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String
Private wsITMCODE As String
Private wsRcd As String
Private wsKeyIn As String

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    waResult.ReDim 0, -1, GITMCODE, GITMID
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
    wlItmID = 0
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    
    FocusMe cboITMCODE
    
End Sub

Private Function Chk_cboItmCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboItmCode = False
    
    If Trim(cboITMCODE.Text) = "" Then
        Chk_cboItmCode = True
        Exit Function
    End If
        
    If Chk_ItmCode(cboITMCODE, wsStatus) = True Then
        
        If wsStatus = "2" Then
            gsMsg = "物料已存在但已無效!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboITMCODE.SetFocus
            Exit Function
        End If
    Else
        gsMsg = "物料不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboITMCODE.SetFocus
        Exit Function
    End If
    
    Chk_cboItmCode = True
End Function

Private Function Chk_cboWhsCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboWhsCode = False
    
    If Trim(cboWhsCode.Text) = "" Then
        Chk_cboWhsCode = True
        Exit Function
    End If
        
    If Chk_WhsCode(cboWhsCode) = False Then
        
        gsMsg = "倉庫不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWhsCode.SetFocus
        Exit Function
    End If
    
    Chk_cboWhsCode = True
End Function
Private Function Chk_WhsCode(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT WhsDesc FROM mstWarehouse WHERE WhsCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And WhsStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_WhsCode = True
    Else
        Chk_WhsCode = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function


Private Function Chk_KeyFld() As Boolean
    
Chk_KeyFld = False

If Trim(cboITMCODE.Text) = "" And Trim(cboWhsCode.Text) = "" Then
    Exit Function
End If


If Trim(cboITMCODE.Text) <> "" Then
If Chk_cboItmCode = False Then Exit Function
wsKeyIn = "1"
wsRcd = cboITMCODE.Text
End If

If Trim(cboWhsCode.Text) <> "" Then
If Chk_cboWhsCode = False Then Exit Function
wsKeyIn = "2"
wsRcd = cboWhsCode.Text
End If


If Trim(cboITMCODE.Text) <> "" And Trim(cboWhsCode.Text) <> "" Then
wsKeyIn = "3"
wsRcd = cboITMCODE.Text & cboWhsCode.Text
End If
    
    
Chk_KeyFld = True
End Function

Private Sub Ini_Scr_AfrKey()
    Call LoadRecord
    
    wiAction = CorRec
    If RowLock(wsConnTime, wsKeyType, wsRcd, wsFormID, wsUsrId) = False Then
        gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
        MsgBox gsMsg, vbOKOnly, gsTitle
        tblDetail.ReBind
    End If
    
    Call SetButtonStatus("AfrKeyEdit")
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    Call SetFieldStatus("AfrKey")
    
    If tblDetail.Enabled = True Then
        
        tblDetail.SetFocus
    End If
   
End Sub

Private Sub cboItmCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboITMCODE
  
    wsSQL = "SELECT ItmCode, ItmBarCode, ItmChiName "
    wsSQL = wsSQL & " FROM MstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboITMCODE.SelLength > 0, "", Set_Quote(cboITMCODE.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmStatus <> '2' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    Call Ini_Combo(3, wsSQL, cboITMCODE.Left, cboITMCODE.Top + cboITMCODE.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCode_GotFocus()
    FocusMe cboITMCODE
End Sub

Private Sub cboItmCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboITMCODE, 13, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboItmCode() = False Then Exit Sub
        
        cboWhsCode.SetFocus
        
    End If
End Sub

Private Sub cboItmCode_LostFocus()
    FocusMe cboITMCODE, True
End Sub
Private Sub cboWhsCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWhsCode
  
    wsSQL = "SELECT WhsCode, WhsDesc "
    wsSQL = wsSQL & " FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsCode LIKE '%" & IIf(cboWhsCode.SelLength > 0, "", Set_Quote(cboWhsCode.Text)) & "%' "
    wsSQL = wsSQL & " AND WhsStatus <> '2' "
    wsSQL = wsSQL & " ORDER BY WhsCode "
    Call Ini_Combo(2, wsSQL, cboWhsCode.Left, cboWhsCode.Top + cboWhsCode.Height, tblCommon, wsFormID, "TBLWhsCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWhsCode_GotFocus()
    FocusMe cboWhsCode
End Sub

Private Sub cboWhsCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboWhsCode() = False Then Exit Sub
        
        If Chk_KeyFld = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If
End Sub

Private Sub cboWhsCode_LostFocus()
    FocusMe cboWhsCode, True
End Sub
Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        Set wcCombo = cboITMCODE
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
        
        'Case vbKeyF9
       
        '    If tbrProcess.Buttons(tcFind).Enabled = True Then Call cmdFind
            
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
    Call Ini_Data
  
    MousePointer = vbDefault
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
 If gsLangID = "1" Then
    
    wsSQL = "SELECT ItmID, ItmEngName ItmName "
    wsSQL = wsSQL & "FROM MstItem "
    wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND ItmCode='" & Set_Quote(cboITMCODE) & "' "
Else
    wsSQL = "SELECT ItmID, ItmChiName ItmName "
    wsSQL = wsSQL & "FROM MstItem "
    wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND ItmCode='" & Set_Quote(cboITMCODE) & "' "

End If

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        lblDspItmName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    Else
        wlItmID = ReadRs(rsRcd, "ItmID")
        lblDspItmName = ReadRs(rsRcd, "ItmName")
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
  
  
 Call LoadGridRecord
 
 LoadRecord = True
 
End Function

Private Function LoadGridRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    
    LoadGridRecord = False
    
    wsSQL = "SELECT WhsItemItmID, ItmCode, ItmChiName, WhsItemWhsCode, WhsItemBinNo, "
    wsSQL = wsSQL & "WhsItemAveCost, WhsItemLPOCost,WhsItemLPrice "
    wsSQL = wsSQL & "FROM MstItem, mstWhsItem "
    wsSQL = wsSQL & "WHERE ItmStatus =  '1' "
    wsSQL = wsSQL & "AND WhsItemStatus = '1' "
    wsSQL = wsSQL & "AND WhsItemItmID = ItmID "
    If wlItmID > 0 Then
    wsSQL = wsSQL & "AND WhsItemItmID = " & wlItmID & " "
    End If
    If Trim(cboWhsCode.Text) <> "" Then
    wsSQL = wsSQL & "AND WhsItemWhsCode = '" & Set_Quote(cboWhsCode.Text) & "' "
    End If
    wsSQL = wsSQL & "ORDER BY ItmCode "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, GITMCODE, GITMID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), GITMCODE) = ReadRs(rsRcd, "ItmCode")
             waResult(.UpperBound(1), GITEMNAME) = ReadRs(rsRcd, "ItmChiName")
             waResult(.UpperBound(1), GWHSCODE) = ReadRs(rsRcd, "WhsItemWhsCode")
             waResult(.UpperBound(1), GLOTNO) = ReadRs(rsRcd, "WhsItemBinNo")
             waResult(.UpperBound(1), GLEAD) = Format(To_Value(ReadRs(rsRcd, "WhsItemLead")), gsQtyFmt)
             waResult(.UpperBound(1), GAVGCOST) = Format(To_Value(ReadRs(rsRcd, "WhsItemAveCost")), gsAmtFmt)
             waResult(.UpperBound(1), GPOCOST) = Format(To_Value(ReadRs(rsRcd, "WhsItemLPOCost")), gsAmtFmt)
             waResult(.UpperBound(1), GLPRICE) = Format(To_Value(ReadRs(rsRcd, "WhsItemLPrice")), gsAmtFmt)
             waResult(.UpperBound(1), GITMID) = To_Value(ReadRs(rsRcd, "WhsItemItmID"))
             rsRcd.MoveNext
         Loop
    End With
    
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    Set rsRcd = Nothing
    
    LoadGridRecord = True
    
    
End Function



Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblItmCode.Caption = Get_Caption(waScrItm, "ITMCODE")
    lblItemChiName.Caption = Get_Caption(waScrItm, "ITEMNAME")
    lblWhsCode.Caption = Get_Caption(waScrItm, "WHSCODE")
    
    With tblDetail
        .Columns(GITMCODE).Caption = Get_Caption(waScrItm, "GITMCODE")
        .Columns(GWHSCODE).Caption = Get_Caption(waScrItm, "GWHSCODE")
        .Columns(GITEMNAME).Caption = Get_Caption(waScrItm, "GITEMNAME")
        .Columns(GLOTNO).Caption = Get_Caption(waScrItm, "GLOTNO")
        .Columns(GLEAD).Caption = Get_Caption(waScrItm, "GLEAD")
        .Columns(GAVGCOST).Caption = Get_Caption(waScrItm, "GAVGCOST")
        .Columns(GPOCOST).Caption = Get_Caption(waScrItm, "GLPOCOST")
        .Columns(GLPRICE).Caption = Get_Caption(waScrItm, "GLPRICE")
        
    End With
    
    
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "IWADD")
    wsActNam(2) = Get_Caption(waScrItm, "IWEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "IWDELETE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 7305
        Me.Width = 9915
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
    Set frmIW001 = Nothing
    
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
    Dim wsCtlPrd As String
    Dim wsSts As String
    Dim bResult As Boolean
    Dim i As Integer
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, wsRcd, wsFormID) Then
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
        adcmdSave.CommandText = "USP_IW001"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GITMCODE)) <> "" And Trim(waResult(wiCtr, GWHSCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wsKeyIn)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, GITMID))
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, GWHSCODE))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, GLOTNO))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, GLEAD))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, GAVGCOST))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, GPOCOST))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, GLPRICE))
                Call SetSPPara(adcmdSave, 9, wiCtr)
                Call SetSPPara(adcmdSave, 10, gsUserID)
                Call SetSPPara(adcmdSave, 11, wsGenDte)
                
                adcmdSave.Execute
                wlKey = GetSPPara(adcmdSave, 12)
            End If
        Next
    End If
    cnCon.CommitTrans
    
    
    If wiAction = CorRec And Trim(wlKey) <> 0 Then
        gsMsg = "已儲存!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    Else
        gsMsg = "儲存失敗!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    
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
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, GITMCODE)) <> "" And Trim(waResult(wlCtr, GWHSCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有設定倉庫物料價!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
        tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
        tblDetail.FirstRow = tblDetail.Row
        tblDetail.Col = GITMCODE
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

    Dim newForm As New frmIW001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()

    Dim newForm As New frmIW001
    
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
    wsFormID = "IW001"

End Sub

Private Sub cmdCancel()
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboITMCODE.SetFocus
End Sub

Private Sub cmdFind()
   ' Call OpenPromptForm
End Sub

Public Property Let ITMCODE(SITMCODE As String)
    wsITMCODE = SITMCODE
End Property

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
        Case tcFind
            Call cmdFind
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
    '    .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = GITMCODE To GDummy
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case GITMCODE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                
                Case GWHSCODE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case GITEMNAME
                    .Columns(wiCtr).Width = 2400
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                    
                Case GLOTNO
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 20
                    
                    
                Case GLEAD
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Alignment = dbgRight
                    
                Case GAVGCOST
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    
                Case GPOCOST
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    
                Case GLPRICE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                
                Case GITMID
                    .Columns(wiCtr).Width = 0
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                Case GDummy
                    .Columns(wiCtr).Width = 10
                    .Columns(wiCtr).DataWidth = 0
                    .Columns(wiCtr).Locked = False
                    
                End Select
        Next
     '   .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
    End With

End Sub


Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wlItmID As Long
    Dim wsItemName As String
    
    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
    
    With tblDetail
        Select Case ColIndex
            Case GITMCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdITMCODE(.Columns(ColIndex).Text, wlItmID, wsItemName) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(GITEMNAME).Text = wsItemName
                .Columns(GLOTNO).Text = ""
                .Columns(GLEAD).Text = Format("0", gsQtyFmt)
                .Columns(GAVGCOST).Text = Format("0", gsAmtFmt)
                .Columns(GPOCOST).Text = Format("0", gsAmtFmt)
                .Columns(GLPRICE).Text = Format("0", gsAmtFmt)
                .Columns(GITMID).Text = wlItmID
                
            Case GWHSCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
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
    Dim wsSQL As String
    Dim wiTop As Long
    
    On Error GoTo tblDetail_ButtonClick_Err

    With tblDetail
        Select Case ColIndex
            Case GITMCODE
            
                If gsLangID = 1 Then
                wsSQL = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME FROM mstITEM "
                wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(cboITMCODE.Text) & "%' "
                wsSQL = wsSQL & " ORDER BY ITMCODE "
                Else
                wsSQL = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME FROM mstITEM "
                wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(cboITMCODE.Text) & "%' "
                wsSQL = wsSQL & " ORDER BY ITMCODE "
                End If
                
                
                Call Ini_Combo(3, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case GWHSCODE
            
                wsSQL = "SELECT WhsCode, WhsDesc FROM MstWarehouse "
                wsSQL = wsSQL & " WHERE WhsStatus <> '2' AND WhsCode LIKE '%" & Set_Quote(cboWhsCode.Text) & "%' "
                wsSQL = wsSQL & " ORDER BY WhsCode"
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
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
                
             If .Col <> GLPRICE Then
                KeyCode = vbDefault
                .Col = .Col + 1
             Else
                KeyCode = vbKeyDown
                .Col = GITMCODE
             End If
              
        Case vbKeyLeft
             KeyCode = vbDefault
             If .Col <> GITMCODE Then
                   .Col = .Col - 1
             End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> GLPRICE Then
                  .Col = .Col + 1
            End If
            
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub



Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
        
        Case GLEAD
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case GAVGCOST, GPOCOST, GLPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
    End Select
    
End Sub


Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = GITMCODE
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case GITMCODE
                    Call Chk_grdITMCODE(.Columns(GITMCODE).Text, 0, "")
                    
                Case GWHSCODE
                    Call Chk_grdWhsCode(.Columns(GWHSCODE).Text)
                    
                    
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub
Private Function Chk_grdITMCODE(inItmCode As String, OutID, outItmName As String) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset

    If wsKeyIn <> "2" And cboITMCODE <> inItmCode Then
        gsMsg = "不同物料名!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
    End If
    
    wsSQL = "SELECT ItmID, ItmChiName FROM MstItem "
    wsSQL = wsSQL & " WHERE ItmCode = '" & Set_Quote(inItmCode) & "' And ItmStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        outItmName = ReadRs(rsRcd, "ItmChiName")
        OutID = ReadRs(rsRcd, "ItmID")
        
        Chk_grdITMCODE = True
    Else
        outItmName = ""
        OutID = 0
        gsMsg = "沒有此物料名!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_grdWhsCode(inWhsCode As String) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset

    If wsKeyIn <> "1" And cboWhsCode <> inWhsCode Then
        gsMsg = "不同倉庫!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdWhsCode = False
    End If
    
    wsSQL = "SELECT WhsDesc FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsCode = '" & Set_Quote(inWhsCode) & "' And WhsStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_grdWhsCode = True
    Else
        gsMsg = "沒有此倉庫!"
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
                If Trim(.Columns(GITMCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, GITMCODE)) = "" And _
                   Trim(waResult(inRow, GWHSCODE)) = "" And _
                   Trim(waResult(inRow, GITEMNAME)) = "" And _
                   Trim(waResult(inRow, GLOTNO)) = "" And _
                   Trim(waResult(inRow, GLEAD)) = "" And _
                   Trim(waResult(inRow, GAVGCOST)) = "" And _
                   Trim(waResult(inRow, GPOCOST)) = "" And _
                   Trim(waResult(inRow, GLPRICE)) = "" Then
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
        
        If Chk_grdITMCODE(waResult(LastRow, GITMCODE), 0, "") = False Then
            .Col = GITMCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdWhsCode(waResult(LastRow, GWHSCODE)) = False Then
                .Col = GWHSCODE
                .Row = LastRow
                Exit Function
        End If
        
        
    End With
    
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check chk_GrdRow"
    
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
        
            Me.cboITMCODE.Enabled = False
            Me.cboWhsCode.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboITMCODE.Enabled = True
            Me.cboWhsCode.Enabled = True
            
       
       Case "AfrActEdit"
       
            Me.cboITMCODE.Enabled = True
            Me.cboWhsCode.Enabled = True
            
        
        Case "AfrKey"
            Me.cboITMCODE.Enabled = False
            Me.cboWhsCode.Enabled = False
            
        If wiAction = CorRec Then
                tblDetail.Enabled = True
        End If
             
    End Select
End Sub


Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(GITMCODE)
    wsCurRecLn = tblDetail.Columns(GWHSCODE)
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, GITMCODE) And _
                  wsCurRecLn = waResult(wlCtr, GWHSCODE) Then
                  gsMsg = "物料名和倉庫已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function


Private Sub Ini_Data()
    If wsITMCODE <> "" Then
        cboITMCODE.Text = wsITMCODE
        If cboITMCODE.Enabled = True Then
            SendKeys "{ENTER}"
        End If
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
