VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmIP001 
   Caption         =   "書本對換價"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "IP001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboITMCODE 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   3495
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "IP001.frx":08CA
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
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
            Picture         =   "IP001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IP001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   4935
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Customer Pricing"
      TabPicture(0)   =   "IP001.frx":6C6D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tblCusItem"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vendor Pricing"
      TabPicture(1)   =   "IP001.frx":6C89
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tblDetail"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TrueDBGrid60.TDBGrid tblCusItem 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "IP001.frx":6CA5
         TabIndex        =   6
         Top             =   120
         Width           =   9495
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   4335
         Left            =   -74880
         OleObjectBlob   =   "IP001.frx":D411
         TabIndex        =   7
         Top             =   120
         Width           =   9495
      End
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
   Begin VB.Label lblPrice 
      Caption         =   "UOMCODE"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDspPrice 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4080
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblCurr 
      Caption         =   "UOMCODE"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblDspCurr 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDspUOMCode 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   6360
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblDspItmName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   8055
   End
   Begin VB.Label lblVdrItemChiName 
      Caption         =   "VDRITEMCHINAME"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label lblUOMCode 
      Caption         =   "UOMCODE"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   900
      Width           =   975
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
      Width           =   1455
   End
   Begin VB.Menu mnuCPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuCPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuVPopUp 
      Caption         =   "PoP Up"
      Visible         =   0   'False
      Begin VB.Menu mnuVPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmIP001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Private waCusResult As New XArrayDB
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private waPopUpSub As New XArrayDB

Private wcCombo As Control

Private wsOldCusNo As String
Private wsOldCurCd As String
Private wsOldShipCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String

Private Const VDRCODE = 0
Private Const VDRNAME = 1
Private Const VDRCURR = 2
Private Const Price = 3
Private Const DISCOUNT = 4
Private Const CNVFACTOR = 5
Private Const UOMCODE = 6
Private Const COST = 7
Private Const PRICEL = 8
Private Const COSTL = 9
Private Const VDRID = 10

Private Const CusCode = 0
Private Const CUSNAME = 1
Private Const CUSCURR = 2
Private Const CUSPRICE = 3
Private Const CUSCNVFACTOR = 4
Private Const CUSUOMCODE = 5
Private Const CUSPRICEL = 6
Private Const CUSID = 7

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
Private Const wsKeyType = "MstVdrItem"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String
Private wsITMCODE As String
Private wsItmExcr As String

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    waResult.ReDim 0, -1, VDRCODE, VDRID
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    waCusResult.ReDim 0, -1, CusCode, CUSID
    Set tblCusItem.Array = waCusResult
    tblCusItem.ReBind
    tblCusItem.Bookmark = 0
    
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
    wsItmExcr = "0"
    
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    
    FocusMe cboITMCODE
    
    tabDetailInfo.TabVisible(0) = False
    tabDetailInfo.Tab = 1
    
End Sub

Private Function Chk_cboItmCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboItmCode = False
    
    If Trim(cboITMCODE.Text) = "" Then
        gsMsg = "必需輸入物料號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboITMCODE.SetFocus
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



Private Sub Ini_Scr_AfrKey()

If LoadRecord = False Then Exit Sub

    
    wiAction = CorRec
    If RowLock(wsConnTime, wsKeyType, cboITMCODE, wsFormID, wsUsrId) = False Then
        gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
        MsgBox gsMsg, vbOKOnly, gsTitle
        tblDetail.ReBind
        tblCusItem.ReBind
    End If
    
    Call SetButtonStatus("AfrKeyEdit")
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    Call SetFieldStatus("AfrKey")
    
    If tabDetailInfo.Tab = 0 Then
        If tblCusItem.Enabled = True Then
            tblCusItem.SetFocus
        End If
    Else
        If tblDetail.Enabled = True Then
            tblDetail.SetFocus
        End If
    End If
End Sub

Private Sub cboItmCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboITMCODE
  
    wsSQL = "SELECT ItmCode, ItmChiName "
    wsSQL = wsSQL & " FROM MstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboITMCODE.SelLength > 0, "", Set_Quote(cboITMCODE.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmStatus <> '2' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    Call Ini_Combo(2, wsSQL, cboITMCODE.Left, cboITMCODE.Top + cboITMCODE.Height, tblCommon, "IP001", "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCode_GotFocus()
    FocusMe cboITMCODE
End Sub

Private Sub cboItmCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboITMCODE, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboItmCode() = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If
End Sub

Private Sub cboItmCode_LostFocus()
    FocusMe cboITMCODE, True
End Sub

Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        Set wcCombo = cboITMCODE
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
   Select Case KeyCode
        
        Case vbKeyPageDown
            KeyCode = 0
            If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab + 1
                Exit Sub
            End If
        Case vbKeyPageUp
            KeyCode = 0
            If tabDetailInfo.Tab > 0 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab - 1
                Exit Sub
            End If
        
        Case vbKeyF6
            Call cmdOpen
        
        
        'Case vbKeyF2
        '    If wiAction = DefaultPage Then Call cmdNew
            
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
       
        
        'Case vbKeyF3
        '    If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF9
        
        If tbrProcess.Buttons(tcFind).Enabled = True Then
            Call cmdFind
        End If
        
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
    Call Ini_CusGrid
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
    
    wsSQL = "SELECT ItmID, ItmEngName ItmName, ItmUomCode, ItmUnitPrice, ItmCurr "
    wsSQL = wsSQL & "FROM MstItem "
    wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND ItmCode='" & Set_Quote(cboITMCODE) & "' "
Else
    wsSQL = "SELECT ItmID, ItmChiName ItmName, ItmUomCode, ItmUnitPrice, ItmCurr "
    wsSQL = wsSQL & "FROM MstItem "
    wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND ItmCode='" & Set_Quote(cboITMCODE) & "' "

End If

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        lblDspItmName = ""
        lblDspUOMCode = ""
        lblDspCurr = ""
        lblDspPrice = ""
        
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    Else
        wlItmID = ReadRs(rsRcd, "ItmID")
        lblDspItmName = ReadRs(rsRcd, "ItmName")
        lblDspUOMCode = ReadRs(rsRcd, "ItmUOMCode")
        lblDspCurr = ReadRs(rsRcd, "ItmCurr")
        lblDspPrice = Format(To_Value(ReadRs(rsRcd, "ItmUnitPrice")), gsUprFmt)
        
                If getExcPRate(lblDspCurr.Caption, gsSystemDate, wsItmExcr, "") = False Then
                    gsMsg = "沒有此貨幣!"
                    wsItmExcr = "1"
                End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
  
 Call LoadCusRecord
 Call LoadVdrRecord
 
 LoadRecord = True
 
End Function

Private Function LoadVdrRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    
    LoadVdrRecord = False
    
    If gsLangID = "1" Then
        wsSQL = "SELECT VdrCode, VdrItemEngName VdrItmName, VdrItemUOMCODE, "
        wsSQL = wsSQL & "VdrItemCurr, VdrItemPrice,VdrItemPricel, VdrItemCnvFactor, VdrItemDiscount, "
        wsSQL = wsSQL & "VdrItemCost, VdrItemCostl, VdrItemID, VdrItemVdrID, VdrName "
        wsSQL = wsSQL & "FROM MstItem, MstVdrItem, MstVendor "
        wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND VdrItemStatus = '1' "
        wsSQL = wsSQL & "AND VdrItemItmID = ItmID AND VdrItemItmID = " & wlItmID & " "
        wsSQL = wsSQL & "AND VdrItemVdrID = VdrID "
        wsSQL = wsSQL & "ORDER BY VdrCode "
    Else
        wsSQL = "SELECT VdrCode, VdrItemChiName VdrItmName, VdrItemUOMCODE, "
        wsSQL = wsSQL & "VdrItemCurr, VdrItemPrice, VdrItemPricel, VdrItemCnvFactor, VdrItemDiscount, "
        wsSQL = wsSQL & "VdrItemCost, VdrItemCostl, VdrItemID, VdrItemVdrID, VdrName "
        wsSQL = wsSQL & "FROM MstItem, MstVdrItem, MstVendor "
        wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND VdrItemStatus = '1' "
        wsSQL = wsSQL & "AND VdrItemItmID = ItmID AND VdrItemItmID = " & wlItmID & " "
        wsSQL = wsSQL & "AND VdrItemVdrID = VdrID "
        wsSQL = wsSQL & "ORDER BY VdrCode "
    End If

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, VDRCODE, VDRID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), VDRCODE) = ReadRs(rsRcd, "VdrCode")
             waResult(.UpperBound(1), VDRNAME) = ReadRs(rsRcd, "VdrName")
             waResult(.UpperBound(1), VDRCURR) = ReadRs(rsRcd, "VdrItemCurr")
             waResult(.UpperBound(1), Price) = Format(To_Value(ReadRs(rsRcd, "VdrItemPrice")), gsUprFmt)
             waResult(.UpperBound(1), PRICEL) = To_Value(ReadRs(rsRcd, "VdrItemPricel"))
             waResult(.UpperBound(1), CNVFACTOR) = Format(To_Value(ReadRs(rsRcd, "VdrItemCnvFactor")), gsAmtFmt)
             waResult(.UpperBound(1), DISCOUNT) = Format(To_Value(ReadRs(rsRcd, "VdrItemDiscount")), gsUprFmt)
             waResult(.UpperBound(1), COST) = Format(To_Value(ReadRs(rsRcd, "VdrItemCost")), gsAmtFmt)
             waResult(.UpperBound(1), COSTL) = Format(To_Value(ReadRs(rsRcd, "VdrItemCostl")), gsAmtFmt)
             waResult(.UpperBound(1), UOMCODE) = ReadRs(rsRcd, "VdrItemUOMCODE")
             waResult(.UpperBound(1), VDRID) = ReadRs(rsRcd, "VdrItemVdrID")
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    LoadVdrRecord = True
End Function


Private Function LoadCusRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    
    LoadCusRecord = False
    
   If gsLangID = "1" Then
        wsSQL = "SELECT CusCode, CusItemEngName CusItmName, CusItemUOMCODE, "
        wsSQL = wsSQL & "CusItemCurr, CusItemPrice, CusItemPriceL, CusItemCnvFactor, "
        wsSQL = wsSQL & "CusItemID, CusItemCusID, CusName "
        wsSQL = wsSQL & "FROM MstItem, MstCusItem, MstCustomer "
        wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND CusItemStatus = '1' "
        wsSQL = wsSQL & "AND CusItemItmID = ItmID AND CusItemItmID = " & wlItmID & " "
        wsSQL = wsSQL & "AND CusItemCusID = CusID "
        wsSQL = wsSQL & "ORDER BY CusCode "
    Else
        wsSQL = "SELECT CusCode, CusItemChiName CusItmName, CusItemUOMCODE, "
        wsSQL = wsSQL & "CusItemCurr, CusItemPrice, CusItemPriceL, CusItemCnvFactor, "
        wsSQL = wsSQL & "CusItemID, CusItemCusID, CusName "
        wsSQL = wsSQL & "FROM MstItem, MstCusItem, MstCustomer "
        wsSQL = wsSQL & "WHERE ItmStatus =  '1' AND CusItemStatus = '1' "
        wsSQL = wsSQL & "AND CusItemItmID = ItmID AND CusItemItmID = " & wlItmID & " "
        wsSQL = wsSQL & "AND CusItemCusID = CusID "
        wsSQL = wsSQL & "ORDER BY CusCode "
    End If

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
   rsRcd.MoveFirst
    With waCusResult
         .ReDim 0, -1, CusCode, CUSID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waCusResult(.UpperBound(1), CusCode) = ReadRs(rsRcd, "CusCode")
             waCusResult(.UpperBound(1), CUSNAME) = ReadRs(rsRcd, "CusName")
             waCusResult(.UpperBound(1), CUSCURR) = ReadRs(rsRcd, "CusItemCurr")
             waCusResult(.UpperBound(1), CUSPRICE) = Format(To_Value(ReadRs(rsRcd, "CusItemPrice")), gsAmtFmt)
             waCusResult(.UpperBound(1), CUSPRICEL) = Format(To_Value(ReadRs(rsRcd, "CusItemPriceL")), gsAmtFmt)
             waCusResult(.UpperBound(1), CUSCNVFACTOR) = Format(To_Value(ReadRs(rsRcd, "CUsItemCnvFactor")), gsAmtFmt)
             waCusResult(.UpperBound(1), CUSUOMCODE) = ReadRs(rsRcd, "CusItemUOMCODE")
             waCusResult(.UpperBound(1), CUSID) = ReadRs(rsRcd, "CusItemCusID")
             rsRcd.MoveNext
         Loop
    End With
    tblCusItem.ReBind
    tblCusItem.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    
    LoadCusRecord = True
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblItmCode.Caption = Get_Caption(waScrItm, "ITEMCODE")
    lblVdrItemChiName.Caption = Get_Caption(waScrItm, "ITEMNAME")
    lblUOMCode.Caption = Get_Caption(waScrItm, "UOMCODE")
    lblCurr.Caption = Get_Caption(waScrItm, "ITMCURR")
    lblPrice.Caption = Get_Caption(waScrItm, "ITMPRICE")
    
    With tblDetail
        .Columns(VDRCODE).Caption = Get_Caption(waScrItm, "VDRCODE")
        .Columns(VDRNAME).Caption = Get_Caption(waScrItm, "VDRNAME")
        .Columns(VDRCURR).Caption = Get_Caption(waScrItm, "VDRCURR")
        .Columns(Price).Caption = Get_Caption(waScrItm, "PRICE")
        .Columns(CNVFACTOR).Caption = Get_Caption(waScrItm, "CNVFACTOR")
        .Columns(DISCOUNT).Caption = Get_Caption(waScrItm, "DISCOUNT")
        .Columns(COST).Caption = Get_Caption(waScrItm, "COST")
        .Columns(COSTL).Caption = Get_Caption(waScrItm, "COSTL")
        .Columns(UOMCODE).Caption = Get_Caption(waScrItm, "VDRUOMCODE")
        .Columns(VDRID).Caption = Get_Caption(waScrItm, "VDRID")
    End With
    
    With tblCusItem
        .Columns(CusCode).Caption = Get_Caption(waScrItm, "CUSCODE")
        .Columns(CUSNAME).Caption = Get_Caption(waScrItm, "CUSNAME")
        .Columns(CUSCURR).Caption = Get_Caption(waScrItm, "CUSCURR")
        .Columns(CUSPRICE).Caption = Get_Caption(waScrItm, "CUSPRICE")
        .Columns(CUSPRICEL).Caption = Get_Caption(waScrItm, "CUSPRICEL")
        .Columns(CUSCNVFACTOR).Caption = Get_Caption(waScrItm, "CUSCNVFACTOR")
        .Columns(CUSUOMCODE).Caption = Get_Caption(waScrItm, "CUSUOMCODE")
        .Columns(CUSID).Caption = Get_Caption(waScrItm, "CUSID")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "IPADD")
    wsActNam(2) = Get_Caption(waScrItm, "IPEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "IPDELETE")
    
    Call Ini_PopMenu(mnuCPopUpSub, "POPUP", waPopUpSub)
    Call Ini_PopMenu(mnuVPopUpSub, "POPUP", waPopUpSub)
    
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
    Set waCusResult = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waPopUpSub = Nothing
    Set frmIP001 = Nothing
End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If tblDetail.Enabled Then
            tblDetail.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
    
        If Me.tblCusItem.Enabled Then
            tblCusItem.SetFocus
        End If
    End If
End Sub

Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        wcCombo.Text = tblCommon.Columns(0).Text
    ElseIf wcCombo.Name = tblCusItem.Name Then
        tblCusItem.EditActive = True
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
        ElseIf wcCombo.Name = tblCusItem.Name Then
            tblCusItem.EditActive = True
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboITMCODE, wsFormID) Then
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
        adcmdSave.CommandText = "USP_IP001V"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, VDRCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlItmID)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, VDRID))
                Call SetSPPara(adcmdSave, 4, "")
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, UOMCODE))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, VDRCURR))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, Price))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, CNVFACTOR))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, DISCOUNT))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, COST))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, COSTL))
                Call SetSPPara(adcmdSave, 12, wiCtr)
                Call SetSPPara(adcmdSave, 13, gsUserID)
                Call SetSPPara(adcmdSave, 14, wsGenDte)
                
                adcmdSave.Execute
                wlKey = GetSPPara(adcmdSave, 15)
            End If
        Next
    End If
    cnCon.CommitTrans
    
    If wiAction = CorRec And Trim(wlKey) <> 0 Then
        bResult = True
    Else
        bResult = False
    End If
    
    Set adcmdSave = Nothing
    cmdSave = True
    
    wlRowCtr = waCusResult.UpperBound(1)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
      
    If waCusResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_IP001C"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waCusResult.UpperBound(1)
            If Trim(waCusResult(wiCtr, CusCode)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlItmID)
                Call SetSPPara(adcmdSave, 3, waCusResult(wiCtr, CUSID))
                Call SetSPPara(adcmdSave, 4, "")
                Call SetSPPara(adcmdSave, 5, waCusResult(wiCtr, CUSUOMCODE))
                Call SetSPPara(adcmdSave, 6, waCusResult(wiCtr, CUSCURR))
                Call SetSPPara(adcmdSave, 7, waCusResult(wiCtr, CUSPRICE))
                Call SetSPPara(adcmdSave, 8, waCusResult(wiCtr, CUSPRICEL))
                Call SetSPPara(adcmdSave, 9, waCusResult(wiCtr, CUSCNVFACTOR))
                Call SetSPPara(adcmdSave, 10, wiCtr)
                Call SetSPPara(adcmdSave, 11, gsUserID)
                Call SetSPPara(adcmdSave, 12, wsGenDte)
                
                adcmdSave.Execute
                wlKey = GetSPPara(adcmdSave, 13)
            End If
        Next
    End If
    cnCon.CommitTrans
    
    If wiAction = CorRec And Trim(wlKey) <> 0 And bResult = True Then
        gsMsg = "物料價已儲存!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    Else
        gsMsg = "客戶或供應商物料價儲存失敗!"
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
    Dim wlCtr1 As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, VDRCODE)) <> "" Then
                
                wiEmptyGrid = False
                If Chk_VdrGrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
                
                For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                If waResult(wlCtr, VDRCODE) = waResult(wlCtr1, VDRCODE) And _
                  waResult(wlCtr, VDRCURR) = waResult(wlCtr1, VDRCURR) Then
                  gsMsg = "供應商或貨幣已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
                End If
                End If
                Next
                
                
                
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有設定供應商物料價!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
        tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    'If Chk_NoVdrDup(To_Value(tblDetail.Bookmark)) = False Then
    '    tblDetail.FirstRow = tblDetail.Row
    '    tblDetail.Col = VDRCODE
    '    tblDetail.SetFocus
    '    Exit Function
    'End If
    
 '   With waCusResult
 '       For wlCtr = 0 To .UpperBound(1)
 '           If Trim(waCusResult(wlCtr, VDRCODE)) <> "" Then
 '               wiEmptyGrid = False
 '               If Chk_CusGrdRow(wlCtr) = False Then
 '                   tblCusItem.SetFocus
 '                   Exit Function
 '               End If
 '           End If
 '       Next
 '   End With
    
 '   If wiEmptyGrid = True Then
 '       gsMsg = "沒有設定客戶物料價!"
 '       MsgBox gsMsg, vbOKOnly, gsTitle
 '       If tblCusItem.Enabled Then
 '       tblCusItem.SetFocus
 '       End If
 '       Exit Function
 '   End If
 '
 '   If Chk_NoCusDup(To_Value(tblCusItem.Bookmark)) = False Then
 '       tblCusItem.FirstRow = tblCusItem.Row
 '       tblCusItem.Col = CusCode
 '       tblCusItem.SetFocus
 '       Exit Function
 '   End If
    
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cmdNew()

    Dim newForm As New frmIP001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()

    Dim newForm As New frmIP001
    
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
    wsFormID = "IP001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "IP"
End Sub

Private Sub cmdCancel()
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboITMCODE.SetFocus
End Sub

Private Sub cmdFind()
    Call OpenPromptForm
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

Private Sub tblCusItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
        PopupMenu mnuCPopUp
    End If
End Sub

Private Sub tblDetail_BeforeRowColChange(Cancel As Integer)

    On Error GoTo tblDetail_BeforeRowColChange_Err
    With tblDetail
        If Chk_VdrGrdRow(To_Value(.Bookmark)) = False Then
            Cancel = True
            Exit Sub
        End If
    End With
    
    Exit Sub
    
tblDetail_BeforeRowColChange_Err:
    
    MsgBox "Check tblDeiail BeforeRowColChange!"
    Cancel = True

End Sub

Private Sub tblCusItem_BeforeRowColChange(Cancel As Integer)

    On Error GoTo tblCusItem_BeforeRowColChange_Err
    With tblCusItem
        If Chk_CusGrdRow(To_Value(.Bookmark)) = False Then
            Cancel = True
            Exit Sub
        End If
    End With
    
    Exit Sub
    
tblCusItem_BeforeRowColChange_Err:
    
    MsgBox "Check tblDeiail BeforeRowColChange!"
    Cancel = True

End Sub


Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
        PopupMenu mnuVPopUp
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
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = VDRCODE To VDRID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case VDRCODE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case VDRNAME
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 60
                    
                Case VDRCURR
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case Price
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                    .Columns(wiCtr).Locked = True
                    
                Case PRICEL
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                Case CNVFACTOR
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsExrFmt
                    
                Case DISCOUNT
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                    
                Case COST
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                    
                Case COSTL
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                    
                Case UOMCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case VDRID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
       ' .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Sub Ini_CusGrid()
    
    Dim wiCtr As Integer

    With tblCusItem
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
     '   .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = CusCode To CUSID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case CusCode
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case CUSNAME
                    .Columns(wiCtr).Width = 4500
                    .Columns(wiCtr).DataWidth = 60
                    
                Case CUSCURR
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case CUSPRICE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    
                Case CUSPRICEL
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                Case CUSCNVFACTOR
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsExrFmt
                    
                Case CUSUOMCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    
                Case CUSID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
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

Private Sub tblCusItem_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblCusItem
        .Update
    End With

End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsVdrID As String
    Dim wsVdrCurr As String
    
    Dim wsVDRNAME As String
    Dim wsUOMCode As String
    Dim wdPrice As Double
    Dim wsExcRate As String
    Dim wsExcDesc As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    wsExcRate = "0"
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
    
    With tblDetail
        Select Case ColIndex
            Case VDRCODE
                If Not Chk_NoVdrDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdVdrCode(.Columns(ColIndex).Text, wsVdrID, wsVdrCurr, wsVDRNAME, wdPrice) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                'Else
                '    .Columns(VDRID).Text = wsVdrID
                '    If Load_ItemPrice(cboItmCode, wsUOMCode, wdPrice) = False Then
                '        GoTo Tbl_BeforeColUpdate_Err
                '    End If
                End If
                
                .Columns(VDRID).Text = wsVdrID
                .Columns(VDRNAME).Text = wsVDRNAME
                .Columns(VDRCURR).Text = wsVdrCurr
                .Columns(UOMCODE).Text = lblDspUOMCode.Caption
                .Columns(Price).Text = Format(wdPrice, gsUprFmt)
                .Columns(CNVFACTOR).Text = Format("1", gsExrFmt)
                .Columns(DISCOUNT).Text = Format("1", gsUprFmt)
                .Columns(COST).Text = Format(wdPrice, gsAmtFmt)
                
                
                If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                End If
                
                If Trim(.Columns(Price).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(PRICEL).Text = Format(To_Value(.Columns(Price).Text) * To_Value(wsExcRate), gsUprFmt)
                End If
                
                If Trim(.Columns(COST).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(COSTL).Text = Format(To_Value(.Columns(COST).Text) * To_Value(wsExcRate), gsUprFmt)
                End If
                
            Case VDRCURR
                If Not Chk_NoVdrDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdVdrCurr(.Columns(ColIndex).Text, wdPrice) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If To_Value(.Columns(Price).Text) <> wdPrice Then
                
                    .Columns(Price).Text = wdPrice
                
                    If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                    End If
                
                    If Trim(.Columns(Price).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(PRICEL).Text = Format(To_Value(.Columns(Price).Text) * To_Value(wsExcRate), gsUprFmt)
                    End If
                
                    If Trim(.Columns(COST).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(COSTL).Text = Format(To_Value(.Columns(COST).Text) * To_Value(wsExcRate), gsUprFmt)
                    End If
                
                
                
                End If
                
            Case UOMCODE
                If Chk_grdUOMCode(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case CNVFACTOR
            
                    If Chk_grdCnvFactor(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                    
                    If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                        gsMsg = "沒有此貨幣!"
                        MsgBox gsMsg, vbOKOnly, gsTitle
                        GoTo Tbl_BeforeColUpdate_Err
                        Exit Sub
                    End If
                
         
                    If Trim(.Columns(Price).Text) <> "" And Trim(.Columns(DISCOUNT).Text) <> "" Then
                        .Columns(COST).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(DISCOUNT).Text) * To_Value(.Columns(CNVFACTOR).Text), gsAmtFmt)
                    End If
                    
                    If Trim(.Columns(COST).Text) <> "" And wsExcRate <> "0" Then
                        .Columns(COSTL).Text = Format(To_Value(.Columns(COST).Text) * To_Value(wsExcRate), gsAmtFmt)
                    End If
                    
            Case Price
                
                    If Chk_grdPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                    
                    If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                        gsMsg = "沒有此貨幣!"
                        MsgBox gsMsg, vbOKOnly, gsTitle
                        GoTo Tbl_BeforeColUpdate_Err
                        Exit Sub
                    End If
                
                    If Trim(.Columns(Price).Text) <> "" And wsExcRate <> "0" Then
                        .Columns(PRICEL).Text = Format(To_Value(.Columns(Price).Text) * To_Value(wsExcRate), gsUprFmt)
                    End If
                    
                    If Trim(.Columns(Price).Text) <> "" And Trim(.Columns(DISCOUNT).Text) <> "" Then
                        .Columns(COST).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(DISCOUNT).Text) * To_Value(.Columns(CNVFACTOR).Text), gsAmtFmt)
                    End If
                    
                    If Trim(.Columns(COST).Text) <> "" And wsExcRate <> "0" Then
                        .Columns(COSTL).Text = Format(To_Value(.Columns(COST).Text) * To_Value(wsExcRate), gsAmtFmt)
                    End If
                    
                  Case DISCOUNT
                
                    If Chk_grdPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                    
                    If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                        gsMsg = "沒有此貨幣!"
                        MsgBox gsMsg, vbOKOnly, gsTitle
                        GoTo Tbl_BeforeColUpdate_Err
                        Exit Sub
                    End If
                
         
                    If Trim(.Columns(Price).Text) <> "" And Trim(.Columns(DISCOUNT).Text) <> "" Then
                        .Columns(COST).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(DISCOUNT).Text) * To_Value(.Columns(CNVFACTOR).Text), gsAmtFmt)
                    End If
                    
                    If Trim(.Columns(COST).Text) <> "" And wsExcRate <> "0" Then
                        .Columns(COSTL).Text = Format(To_Value(.Columns(COST).Text) * To_Value(wsExcRate), gsAmtFmt)
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

Private Sub tblCusItem_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsCusID As String
    Dim wsCusCurr As String
    
    Dim wsName As String
    Dim wsUOMCode As String
    Dim wdPrice As Double
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wsVdrID As String

    On Error GoTo tblCusItem_BeforeColUpdate_Err
    
    wsExcRate = "0"
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblCusItem.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
    
    With tblCusItem
        Select Case ColIndex
            Case CusCode
                If Not Chk_NoCusDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdCusCode(.Columns(ColIndex).Text, wsCusID, wsCusCurr, wsName) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                Else
                    .Columns(CUSID).Text = wsCusID
                    If Load_ItemPrice(cboITMCODE, wsUOMCode, wdPrice) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                End If
                
                
                .Columns(CUSNAME).Text = wsName
                .Columns(CUSCURR).Text = wsCusCurr
                .Columns(CUSUOMCODE).Text = wsUOMCode

                .Columns(CUSPRICE).Text = Format(wdPrice, gsAmtFmt)
                .Columns(CUSCNVFACTOR).Text = Format("1", gsExrFmt)
                
                If getExcRate(.Columns(CUSCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                End If
                
                If Trim(.Columns(CUSPRICE).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(CUSPRICEL).Text = Format(To_Value(.Columns(CUSPRICE).Text) * To_Value(wsExcRate), gsAmtFmt)
                End If
                
            Case CUSCURR
                If Not Chk_NoCusDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdCusCurr(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case CUSUOMCODE
                If Chk_grdUOMCode(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case CUSCNVFACTOR, CUSPRICE
            
                If ColIndex = CUSCNVFACTOR Then
                    If Chk_grdCnvFactor(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
                End If
                    
                If getExcPRate(.Columns(VDRCURR).Text, gsSystemDate, wsExcRate, wsExcDesc) = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo Tbl_BeforeColUpdate_Err
                    Exit Sub
                End If
                
                If Trim(.Columns(Price).Text) <> "" And wsExcRate <> "0" Then
                    .Columns(CUSPRICEL).Text = Format(To_Value(.Columns(CUSPRICE).Text) * To_Value(wsExcRate), gsAmtFmt)
                End If
            End Select
    End With
    
    Exit Sub
    
Tbl_BeforeColUpdate_Err:
    tblCusItem.Columns(ColIndex).Text = OldValue
    Cancel = True
    Exit Sub

tblCusItem_BeforeColUpdate_Err:
    
    MsgBox "Check tblDeiail BeforeColUpdate!"
    tblCusItem.Columns(ColIndex).Text = OldValue
    Cancel = True
    
End Sub

Private Sub tblDetail_ButtonClick(ByVal ColIndex As Integer)
    Dim wsSQL As String
    Dim wiTop As Long
    
    On Error GoTo tblDetail_ButtonClick_Err

    With tblDetail
        Select Case ColIndex
            Case VDRCODE
                wsSQL = "SELECT VdrCode, VdrName FROM MstVendor "
                wsSQL = wsSQL & " WHERE VdrStatus <> '2' AND VdrCode LIKE '%" & Set_Quote(.Columns(VDRCODE).Text) & "%' "
                wsSQL = wsSQL & " AND VdrInactive = 'N' "
                wsSQL = wsSQL & " ORDER BY VdrCode"
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLVDRCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case UOMCODE
                wsSQL = "SELECT UOMCode, UOMDesc FROM MstUOM "
                wsSQL = wsSQL & " WHERE UOMStatus <> '2' AND UOMCode LIKE '%" & Set_Quote(.Columns(UOMCODE).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY UOMCode"
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLUOMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
            
            Case VDRCURR
                wsSQL = "SELECT DISTINCT ExcCurr FROM MstExchangeRate "
                wsSQL = wsSQL & " WHERE ExcStatus <> '2' AND ExcCurr LIKE '%" & Set_Quote(.Columns(VDRCURR).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY ExcCurr"
                
                Call Ini_Combo(1, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLEXCR", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
        End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDeiail ButtonClick!"
 
    
End Sub

Private Sub tblCusItem_ButtonClick(ByVal ColIndex As Integer)
    Dim wsSQL As String
    Dim wiTop As Long
    
    On Error GoTo tblCusItem_ButtonClick_Err

    With tblCusItem
        Select Case ColIndex
            Case CusCode
                wsSQL = "SELECT CusCode, CusName FROM MstCustomer "
                wsSQL = wsSQL & " WHERE CusStatus <> '2' AND CusCode LIKE '%" & Set_Quote(.Columns(CusCode).Text) & "%' "
                wsSQL = wsSQL & " AND CusInactive = 'N' "
                wsSQL = wsSQL & " ORDER BY CusCode"
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCUSCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblCusItem
                
            Case CUSUOMCODE
                wsSQL = "SELECT UOMCode, UOMDesc FROM MstUOM "
                wsSQL = wsSQL & " WHERE UOMStatus <> '2' AND UOMCode LIKE '%" & Set_Quote(.Columns(CUSUOMCODE).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY UOMCode"
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLUOMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblCusItem
            
            Case CUSCURR
                wsSQL = "SELECT DISTINCT ExcCurr FROM MstExchangeRate "
                wsSQL = wsSQL & " WHERE ExcStatus <> '2' AND ExcCurr LIKE '%" & Set_Quote(.Columns(CUSCURR).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY ExcCurr"
                
                Call Ini_Combo(1, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLEXCR", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblCusItem
                
        End Select
    End With
    
    Exit Sub
    
tblCusItem_ButtonClick_Err:
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
            If IsEmptyVdrRow Then Exit Sub
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
            Select Case .Col
                   
                Case VDRCODE, VDRNAME, VDRCURR, Price, CNVFACTOR, DISCOUNT, UOMCODE
                    KeyCode = vbDefault
                    .Col = .Col + 1
                    
                Case COST
                    KeyCode = vbKeyDown
                    .Col = VDRCODE
                
              End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> VDRCODE Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> UOMCODE Then
                  .Col = .Col + 1
            End If
            
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblCusItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblCusItem_KeyDown_Err
    
    With tblCusItem
        Select Case KeyCode
        Case vbKeyF4        ' CALL COMBO BOX
            KeyCode = vbDefault
            Call tblCusItem_ButtonClick(.Col)
        
        Case vbKeyF5        ' INSERT LINE
            KeyCode = vbDefault
            If .Bookmark = waCusResult.UpperBound(1) Then Exit Sub
            If IsEmptyCusRow Then Exit Sub
            waCusResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
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
            Select Case .Col
                   
                Case CusCode, CUSNAME, CUSCURR, CUSPRICE, CUSCNVFACTOR
                    KeyCode = vbDefault
                    .Col = .Col + 1
               
                Case CUSUOMCODE
                    KeyCode = vbKeyDown
                    .Col = CusCode
                    
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> CusCode Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> CUSUOMCODE Then
                  .Col = .Col + 1
            End If
            
        End Select
    End With

    Exit Sub
    
tblCusItem_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    Select Case tblDetail.Col

        
        Case Price, CNVFACTOR, DISCOUNT
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
    End Select
End Sub

Private Sub tblCusItem_KeyPress(KeyAscii As Integer)
    Select Case tblCusItem.Col
        Case CUSPRICE, CUSCNVFACTOR
            Call Chk_InpNum(KeyAscii, tblCusItem.Text, False, True)
    End Select
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyVdrRow() Then
           .Col = VDRCODE
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case VDRCODE
                    Call Chk_grdVdrCode(.Columns(VDRCODE).Text, .Columns(VDRID).Text, "", "", 0)
                    
                Case UOMCODE
                    Call Chk_grdUOMCode(.Columns(UOMCODE).Text)
                    
                Case VDRCURR
                    Call Chk_grdVdrCurr(.Columns(VDRCURR).Text, 0)
                    
                Case Price
                    Call Chk_grdPrice(.Columns(Price).Text)
                    
                Case CNVFACTOR
                    Call Chk_grdCnvFactor(.Columns(CNVFACTOR).Text)
                
                Case DISCOUNT
                    Call Chk_grdPrice(.Columns(DISCOUNT).Text)
                    
                                   
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Sub tblCusItem_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblCusItem.Name Then Exit Sub
    
    With tblCusItem
        If IsEmptyCusRow() Then
           .Col = VDRCODE
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case CusCode
                    Call Chk_grdCusCode(.Columns(CusCode).Text, .Columns(CUSID).Text, "", "")
                    
                Case CUSUOMCODE
                    Call Chk_grdUOMCode(.Columns(CUSUOMCODE).Text)
                    
                Case CUSCURR
                    Call Chk_grdCusCurr(.Columns(CUSCURR).Text)
                    
                Case CUSPRICE
                    Call Chk_grdPrice(.Columns(CUSPRICE).Text)
                    
                Case CUSCNVFACTOR
                    Call Chk_grdCnvFactor(.Columns(CUSCNVFACTOR).Text)
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblCusItem RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdCusCode(inCusCode As String, outCusID As String, outCusCurr As String, outCusName As String) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset

    wsSQL = "SELECT CusCurr, CusID, CusName FROM MstCustomer"
    wsSQL = wsSQL & " WHERE CusCode = '" & Set_Quote(inCusCode) & "' And CusStatus = '1'"
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outCusCurr = ReadRs(rsDes, "CusCurr")
        outCusID = ReadRs(rsDes, "CusID")
        outCusName = ReadRs(rsDes, "CusName")
        
        Chk_grdCusCode = True
    Else
        outCusCurr = ""
        outCusID = ""
        outCusName = ""
    
        gsMsg = "沒有此客戶!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCusCode = False
    End If
    rsDes.Close
    Set rsDes = Nothing
End Function

Private Function Chk_grdVdrCode(inVdrCode As String, OutVdrID As String, outVdrCurr As String, outVdrName As String, outPrice As Double) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    Dim wsExcRate As String
    Dim wdItmPrice As Double

    wsSQL = "SELECT VdrCurr, VdrID, VdrName FROM MstVendor"
    wsSQL = wsSQL & " WHERE VdrCode = '" & Set_Quote(inVdrCode) & "' And VdrStatus = '1'"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outVdrCurr = ReadRs(rsDes, "VdrCurr")
        OutVdrID = ReadRs(rsDes, "VdrID")
        outVdrName = ReadRs(rsDes, "VdrName")
        
        Chk_grdVdrCode = True
    Else
        outVdrCurr = ""
        OutVdrID = ""
        outVdrName = ""
    
        gsMsg = "沒有此供應商!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdVdrCode = False
        rsDes.Close
        Set rsDes = Nothing
        Exit Function
        
    End If
    rsDes.Close
    Set rsDes = Nothing
    
    wdItmPrice = To_Value(lblDspPrice.Caption)
    outPrice = 0
    
    If UCase(outVdrCurr) <> UCase(lblDspCurr.Caption) Then
    
    If getExcPRate(outVdrCurr, gsSystemDate, wsExcRate, "") = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    Chk_grdVdrCode = False
                    Exit Function
    End If

    
    If To_Value(wsExcRate) <> 0 Then
    outPrice = Format(wdItmPrice * To_Value(wsItmExcr) / To_Value(wsExcRate), gsUprFmt)
    End If
    
    Else
    outPrice = wdItmPrice
    End If
    
End Function

Private Function Chk_grdCusCurr(InCurr As String) As Boolean
    Dim rsExcCurr As New ADODB.Recordset
    Dim Criteria As String
    Dim CtlYr As String
    Dim CtlMon As String
    Dim tmpDte As String
          
    tmpDte = Dsp_Date(gsSystemDate)
    CtlYr = Format(tmpDte, "yyyy")
    CtlMon = Format(tmpDte, "mm")
    
    Criteria = ""
    
    Criteria = Criteria & "SELECT ExcCurr FROM MstExchangeRate "
    Criteria = Criteria & "WHERE EXCCURR = '" & Set_Quote(InCurr) & "' "
    Criteria = Criteria & "AND EXCYR = '" & Set_Quote(CtlYr) & "' "
    Criteria = Criteria & "AND EXCMN = '" & To_Value(CtlMon) & "' "

    rsExcCurr.Open Criteria, cnCon, adOpenStatic, adLockOptimistic
    
    If rsExcCurr.RecordCount > 0 Then
        Chk_grdCusCurr = True
    Else
        gsMsg = "沒有此貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCusCurr = False
    End If
        
    rsExcCurr.Close
    Set rsExcCurr = Nothing
End Function

Private Function Chk_grdVdrCurr(ByVal InCurr As String, ByRef outPrice As Double) As Boolean
    Dim wsExcRate As String
          
    Chk_grdVdrCurr = True
    outPrice = 0
    
    If UCase(InCurr) <> UCase(lblDspCurr.Caption) Then
    
    If getExcPRate(InCurr, gsSystemDate, wsExcRate, "") = False Then
                    gsMsg = "沒有此貨幣!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    Chk_grdVdrCurr = False
                    Exit Function
    End If

    
    If To_Value(wsExcRate) <> 0 Then
    outPrice = Format(To_Value(lblDspPrice.Caption) * To_Value(wsItmExcr) / To_Value(wsExcRate), gsUprFmt)
    End If
    
    Else
    outPrice = To_Value(lblDspPrice.Caption)
    End If
    
End Function

Private Function Chk_grdUOMCode(inUOM As String) As Boolean
    Dim rsUOM As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = ""
    
    wsSQL = wsSQL & "SELECT UOMCode FROM MstUOM "
    wsSQL = wsSQL & "WHERE UOMCode = '" & Set_Quote(inUOM) & "' AND UOMStatus ='1'"

    rsUOM.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsUOM.RecordCount > 0 Then
        Chk_grdUOMCode = True
    Else
        gsMsg = "沒有此單位!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdUOMCode = False
    End If
        
    rsUOM.Close
    Set rsUOM = Nothing
End Function

Private Function Chk_grdCnvFactor(inCode As String) As Boolean
    Chk_grdCnvFactor = True
    
    If Trim(inCode) = "" Then
        gsMsg = "沒有輸入轉換數!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCnvFactor = False
        Exit Function
    End If

    If To_Value(inCode) = 0 Then
        gsMsg = "轉換數不可以為零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCnvFactor = False
        Exit Function
    End If
End Function



Private Function Chk_grdPrice(inCode As String) As Boolean
    Chk_grdPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "沒有輸入價錢!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdPrice = False
        Exit Function
    End If
End Function


Private Function IsEmptyCusRow(Optional inRow) As Boolean

    IsEmptyCusRow = True
    
        
        If IsMissing(inRow) Then
            With tblCusItem
                If Trim(.Columns(CusCode)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waCusResult.UpperBound(1) >= 0 Then
                If Trim(waCusResult(inRow, CusCode)) = "" And _
                   Trim(waCusResult(inRow, CUSNAME)) = "" And _
                   Trim(waCusResult(inRow, CUSCURR)) = "" And _
                   Trim(waCusResult(inRow, CUSPRICE)) = "" And _
                   Trim(waCusResult(inRow, CUSPRICEL)) = "" And _
                   Trim(waCusResult(inRow, CUSCNVFACTOR)) = "" And _
                   Trim(waCusResult(inRow, CUSUOMCODE)) = "" And _
                   Trim(waCusResult(inRow, CUSID)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyCusRow = False
    
End Function
Private Function IsEmptyVdrRow(Optional inRow) As Boolean

    IsEmptyVdrRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(VDRCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, VDRCODE)) = "" And _
                   Trim(waResult(inRow, VDRNAME)) = "" And _
                   Trim(waResult(inRow, VDRCURR)) = "" And _
                   Trim(waResult(inRow, Price)) = "" And _
                   Trim(waResult(inRow, PRICEL)) = "" And _
                   Trim(waResult(inRow, CNVFACTOR)) = "" And _
                   Trim(waResult(inRow, DISCOUNT)) = "" And _
                   Trim(waResult(inRow, COST)) = "" And _
                   Trim(waResult(inRow, COSTL)) = "" And _
                   Trim(waResult(inRow, UOMCODE)) = "" And _
                   Trim(waResult(inRow, VDRID)) = "" Then
                   Exit Function
                End If
            End If
        End If
        
    
    IsEmptyVdrRow = False
    
End Function
Private Function Chk_CusGrdRow(ByVal LastRow As Long) As Boolean
    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    
    Chk_CusGrdRow = False
    
    On Error GoTo Chk_CusGrdRow_Err
    
    
    With tblCusItem
        If To_Value(LastRow) > waCusResult.UpperBound(1) Then
           Chk_CusGrdRow = True
           Exit Function
        End If
        
        If IsEmptyCusRow(To_Value(LastRow)) = True Then
            .Delete
            .Refresh
            .SetFocus
            Chk_CusGrdRow = False
            Exit Function
        End If
        
        If Chk_grdCusCode(waCusResult(LastRow, CusCode), 0, "", "") = False Then
            .Col = CusCode
            Exit Function
        End If
        
        If Chk_grdUOMCode(waCusResult(LastRow, CUSUOMCODE)) = False Then
                .Col = CUSUOMCODE
                Exit Function
        End If
        
        If Chk_grdCusCurr(waCusResult(LastRow, CUSCURR)) = False Then
                .Col = CUSCURR
                Exit Function
        End If
        
        If Chk_grdPrice(waCusResult(LastRow, CUSPRICE)) = False Then
                .Col = CUSPRICE
                Exit Function
        End If
        
        If Chk_grdCnvFactor(waCusResult(LastRow, CUSCNVFACTOR)) = False Then
                .Col = CUSCNVFACTOR
                Exit Function
        End If
        
    End With
        
    Chk_CusGrdRow = True

    Exit Function
    
Chk_CusGrdRow_Err:
    MsgBox "Check Chk_CusGrdRow"
    
End Function
Private Function Chk_VdrGrdRow(ByVal LastRow As Long) As Boolean
    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    
    Chk_VdrGrdRow = False
    
    On Error GoTo Chk_VdrGrdRow_Err
    
    With tblDetail
        If To_Value(LastRow) > waResult.UpperBound(1) Then
           Chk_VdrGrdRow = True
           Exit Function
        End If
        
        If IsEmptyVdrRow(To_Value(LastRow)) = True Then
            .Delete
            .Refresh
            .SetFocus
            Chk_VdrGrdRow = False
            Exit Function
        End If
        
        If Chk_grdVdrCode(waResult(LastRow, VDRCODE), 0, "", "", 0) = False Then
            .Col = VDRCODE
            Exit Function
        End If
        
        If Chk_grdUOMCode(waResult(LastRow, UOMCODE)) = False Then
                .Col = UOMCODE
                Exit Function
        End If
        
        If Chk_grdVdrCurr(waResult(LastRow, VDRCURR), 0) = False Then
                .Col = VDRCURR
                Exit Function
        End If
        
        If Chk_grdPrice(waResult(LastRow, Price)) = False Then
                .Col = Price
                Exit Function
        End If
        
        If Chk_grdCnvFactor(waResult(LastRow, CNVFACTOR)) = False Then
                .Col = CNVFACTOR
                Exit Function
        End If
        
        If Chk_grdPrice(waResult(LastRow, DISCOUNT)) = False Then
                .Col = DISCOUNT
                Exit Function
        End If
        
         
    End With
    
        
    Chk_VdrGrdRow = True

    Exit Function
    
Chk_VdrGrdRow_Err:
    MsgBox "Check Chk_VdrGrdRow"
    
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
            
            Me.tblDetail.Enabled = False
            Me.tblCusItem.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboITMCODE.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboITMCODE.Enabled = True
        
        Case "AfrKey"
            Me.cboITMCODE.Enabled = False
            
            If wiAction = CorRec Then
                Me.tblDetail.Enabled = True
                Me.tblCusItem.Enabled = True
            End If
    End Select
End Sub

Private Function Load_ItemPrice(inItmCode As String, outUOMCode As String, outPrice As Double) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset

    
    wsSQL = "SELECT ItmUOMCode, ItmUnitPrice FROM MstItem"
    wsSQL = wsSQL & " WHERE ItmCode = '" & Set_Quote(inItmCode) & "'"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outUOMCode = ReadRs(rsDes, "ItmUOMCode")
        outPrice = ReadRs(rsDes, "ItmUnitPrice")
    
        
        Load_ItemPrice = True
    Else
        outUOMCode = ""
        outPrice = 0
    
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Load_ItemPrice = False
    End If
    rsDes.Close
    Set rsDes = Nothing
End Function

Private Function Chk_NoVdrDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoVdrDup = False
    
    wsCurRec = tblDetail.Columns(VDRCODE)
    wsCurRecLn = tblDetail.Columns(VDRCURR)
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, VDRCODE) And _
                  wsCurRecLn = waResult(wlCtr, VDRCURR) Then
                  gsMsg = "供應商或貨幣已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoVdrDup = True

End Function

Private Function Chk_NoCusDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoCusDup = False
    
    wsCurRec = tblCusItem.Columns(CusCode)
    wsCurRecLn = tblCusItem.Columns(CUSCURR)
   
        For wlCtr = 0 To waCusResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waCusResult(wlCtr, CusCode) And _
                  wsCurRecLn = waCusResult(wlCtr, CUSCURR) Then
                  gsMsg = "客戶或貨幣已重覆!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoCusDup = True

End Function

Private Sub Ini_Data()
    If wsITMCODE <> "" Then
        cboITMCODE.Text = wsITMCODE
        If cboITMCODE.Enabled = True Then
            SendKeys "{ENTER}"
        End If
    End If
End Sub


Private Sub mnuCPopUpSub_Click(Index As Integer)
    Call Call_CPopUpMenu(waPopUpSub, Index)
End Sub

Private Sub mnuVPopUpSub_Click(Index As Integer)
    Call Call_VPopUpMenu(waPopUpSub, Index)
End Sub

Private Sub Call_CPopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

    Dim wsAct As String
    
    wsAct = inArray(inMnuIdx, 0)
    
    With tblCusItem
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
            
            If .Bookmark = waCusResult.UpperBound(1) Then Exit Sub
            If IsEmptyCusRow Then Exit Sub
            waCusResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
             
    
End Sub

Private Sub Call_VPopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

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
            If IsEmptyVdrRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
             
    
End Sub



Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSQL As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = IIf(gsLangID = "1", "Item Code", "物料編碼")
    vFilterAry(1, 2) = "ItmCode"
    
    vFilterAry(2, 1) = IIf(gsLangID = "1", "Description", "物料名稱")
    vFilterAry(2, 2) = IIf(gsLangID = "1", "ItmEngName", "ItmChiName")
    
    vFilterAry(3, 1) = IIf(gsLangID = "1", "Item Type", "物料分類")
    vFilterAry(3, 2) = "ItmItmTypeCode"
       
    
    ReDim vAry(3, 3)
    vAry(1, 1) = IIf(gsLangID = "1", "Item Code", "物料編碼")
    vAry(1, 2) = "ItmCode"
    vAry(1, 3) = "2500"
    
    vAry(2, 1) = IIf(gsLangID = "1", "Description", "物料名稱")
    vAry(2, 2) = IIf(gsLangID = "1", "ItmEngName", "ItmChiName")
    vAry(2, 3) = "2500"
    
    vAry(3, 1) = IIf(gsLangID = "1", "Item Type", "物料分類")
    vAry(3, 2) = "ItmItmTypeCode"
    vAry(3, 3) = "2500"
        

   wsSQL = "SELECT ItmCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & ", ItmItmTypeCode "
   wsSQL = wsSQL & "FROM MstItem, MstVdrItem, MstVendor "
        
   
    'frmShareSearch.Show vbModal
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE ItmStatus = '1' AND VdrItemStatus = '1' AND VdrItemItmID = ItmID AND VdrItemVdrID = VdrID "
        .sBindOrderSQL = "ORDER BY ItmCode, VdrCode "
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And frmShareSearch.Tag <> cboITMCODE Then
        cboITMCODE = frmShareSearch.Tag
       If cboITMCODE.Enabled = False Then
        LoadRecord
        'txtItmBarCode.Text = ""
        'txtItmCode.SetFocus
       Else
        cboITMCODE.SetFocus
        SendKeys "{Enter}"
       End If
    End If
    Unload frmShareSearch
    
End Sub
