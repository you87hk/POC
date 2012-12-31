VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmAPV0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPV0011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10677.37
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10455
      Begin VB.Label lblDspVdrNo 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label lblDspDocNo 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblVdrNo 
         Caption         =   "Customer Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblDocNo 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   1890
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmAPV0011.frx":0442
      TabIndex        =   0
      Top             =   1560
      Width           =   10455
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":86E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":8FBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":9899
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":9CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":A13D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":A457
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":A8A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":ACFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":B015
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":B32F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":B781
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":C05D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":C385
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":C7D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":CAF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":CE11
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":D265
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":D581
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":D8A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":DBC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPV0011.frx":E015
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   9855
   End
End
Attribute VB_Name = "frmAPV0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private wbErr As Boolean

Private wsVdrNo As String
Private wlDocID As Long
Private wlLastRow As Long

Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String
Private wbUpdFlg As Boolean
Private wsTrnCd As String


Private Const tcRefresh = "Refresh"
Private Const tcExit = "Exit"

Private Const SDOCLINE = 0
Private Const SITMCODE = 1
Private Const SITMNAME = 2
Private Const SITMTYPE = 3
Private Const SUPRICE = 4
Private Const SQTY = 5
Private Const SSOH = 6
Private Const SRECQTY = 7
Private Const SRELQTY = 8
Private Const SDUMMY = 9
Private Const SID = 10



Public Property Let InDocID(inDoc As Long)
     wlDocID = inDoc
End Property


Public Property Let InVdrNo(inDocCus As String)
wsVdrNo = inDocCus
End Property






Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF7
            Call LoadRecord
        
            
        Case vbKeyEscape
            Unload Me
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
    
    waResult.ReDim 0, -1, SDOCLINE, SID
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
    
    
    Select Case wsTrnCd
    Case "PO"
    lblDspDocNo.Caption = Get_TableInfo("POPPOHD", "POHDDOCID = " & wlDocID, "POHDDOCNO")
  '  lblDspRefNo.Caption = Get_TableInfo("POPPOHD", "POHDDOCID = " & wlDocID, "POHDCUSPO")
    
    Case "PV"
    lblDspDocNo.Caption = Get_TableInfo("POPPVHD", "PVHDDOCID = " & wlDocID, "PVHDDOCNO")
  '  lblDspRefNo.Caption = Get_TableInfo("POPPVHD", "PVHDDOCID = " & wlDocID, "PVHDCUSPO")
    
    Case "GR"
    lblDspDocNo.Caption = Get_TableInfo("POPGRHD", "GRHDDOCID = " & wlDocID, "GRHDDOCNO")
    
    Case "PR"
    lblDspDocNo.Caption = Get_TableInfo("POPPRHD", "PRHDDOCID = " & wlDocID, "PRHDDOCNO")
  '  lblDspRefNo.Caption = Get_TableInfo("POPPRHD", "PRHDDOCID = " & wlDocID, "PRHDCUSPO")
    
    End Select
    
    lblDspVdrNo.Caption = wsVdrNo & " - " & Get_TableInfo("MSTVENDOR", "VDRCODE = '" & Set_Quote(wsVdrNo) & "'", "VDRNAME")
    
  '  wbUpdate = False
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    Set frmAPV0011 = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
   ' wsFormID = "APR0011"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
 '   lblRefNo.Caption = Get_Caption(waScrItm, "REFNO")
    lblVdrNo.Caption = Get_Caption(waScrItm, "VdrNo")
    
    With tblDetail
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SSOH).Caption = Get_Caption(waScrItm, "SSOH")
        .Columns(SRELQTY).Caption = Get_Caption(waScrItm, "SRELQTY")
        .Columns(SRECQTY).Caption = Get_Caption(waScrItm, "SRECQTY")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
    End With
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
    
     
 
End Sub





Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
    End With


End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

Dim wsITMID As String
Dim wsITMCODE As String
Dim wsLnType As String
Dim wsDesc As String
Dim wdPrice As Double
Dim wdCost As Double

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
  
       
    
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



Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
       
        Case vbKeyReturn
            Select Case .Col
                Case SRELQTY
                    KeyCode = vbKeyDown
                    .Col = SDOCLINE
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SDOCLINE Then
                   .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SRELQTY
                    KeyCode = vbKeyDown
                    .Col = SDOCLINE
                Case Else
                    KeyCode = vbDefault
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
        
        
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
       
        
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub


Private Function Chk_grdQty(inCode As String) As Boolean
    
    Chk_grdQty = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdQty = False
        Exit Function
    End If

    
    
    
End Function

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SDOCLINE To SID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SDOCLINE
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Width = 500
                Case SITMCODE
                    .Columns(wiCtr).DataWidth = 30
                    .Columns(wiCtr).Width = 2000
                Case SITMNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 2500
                Case SITMTYPE
                   .Columns(wiCtr).Width = 1000
                   .Columns(wiCtr).DataWidth = 10
                Case SUPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                Case SQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SSOH
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SRELQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = False
                Case SRECQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    '.Columns(wiCtr).Locked = False
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
                Case SID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 15
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wdQty As Double
    Dim wsItmCd As String
    
    LoadRecord = False
    Me.MousePointer = vbHourglass
    
    Select Case wsTrnCd
    
    Case "PO"
    
        wsSQL = "SELECT PODTID SID, PODTDOCLINE DOCLINE, PODTITEMID ITEMID, ITMCODE, PODTITEMDESC ITEMDESC, ITMITMTYPECODE, PODTUPRICE UPRICE, PODTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " PODTQTY QTY, PODTRECQTY RECQTY, PODTRELQTY RELQTY "
        wsSQL = wsSQL & " FROM POPPODT, mstITEM "
        wsSQL = wsSQL & " WHERE PODTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND PODTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY PODTDOCLINE "

        
    Case "PV"
    
        wsSQL = "SELECT PVDTID SID, PVDTDOCLINE DOCLINE, PVDTITEMID ITEMID, ITMCODE, PVDTITEMDESC ITEMDESC, ITMITMTYPECODE, PVDTUPRICE UPRICE, PVDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " PVDTQTY QTY, 0 RECQTY, 0 RELQTY "
        wsSQL = wsSQL & " FROM POPPVDT, mstITEM "
        wsSQL = wsSQL & " WHERE PVDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND PVDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY PVDTDOCLINE "
  
    Case "GR"
    
        wsSQL = "SELECT GRDTID SID, GRDTDOCLINE DOCLINE, GRDTITEMID ITEMID, ITMCODE, GRDTITEMDESC ITEMDESC, ITMITMTYPECODE, GRDTUPRICE UPRICE, GRDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " GRDTQTY QTY, 0 RECQTY, 0 RELQTY "
        wsSQL = wsSQL & " FROM POPGRDT, mstITEM "
        wsSQL = wsSQL & " WHERE GRDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND GRDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY GRDTDOCLINE "
  
    Case "PR"
    
        wsSQL = "SELECT PRDTID SID, PRDTDOCLINE DOCLINE, PRDTITEMID ITEMID, ITMCODE, PRDTITEMDESC ITEMDESC, ITMITMTYPECODE, PRDTUPRICE UPRICE, PRDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " PRDTQTY QTY, 0 RECQTY, 0 RELQTY "
        wsSQL = wsSQL & " FROM POPPRDT, mstITEM "
        wsSQL = wsSQL & " WHERE PRDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND PRDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY PRDTDOCLINE "
    
    End Select
     rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    wiCtr = 1
    With waResult
    .ReDim 0, -1, SDOCLINE, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

    wdQty = Get_StockAvailable(To_Value(ReadRs(rsRcd, "ITEMID")), "", "")
    
     .AppendRows
        
        waResult(.UpperBound(1), SDOCLINE) = Format(wiCtr, "000")
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "ITEMDESC")
        waResult(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
        waResult(.UpperBound(1), SUPRICE) = Format(To_Value(ReadRs(rsRcd, "UPRICE")), gsUprFmt)
        waResult(.UpperBound(1), SRECQTY) = IIf(wsTrnCd <> "PO", "", Format(To_Value(ReadRs(rsRcd, "RECQTY")), gsQtyFmt))
        waResult(.UpperBound(1), SRELQTY) = IIf(wsTrnCd <> "PO", "", Format(To_Value(ReadRs(rsRcd, "RELQTY")), gsQtyFmt))
        waResult(.UpperBound(1), SSOH) = Format(wdQty, gsQtyFmt)
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), SID) = To_Value(ReadRs(rsRcd, "SID"))
        wiCtr = wiCtr + 1
        
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



Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 If tbrProcess.Buttons(Button.Key).Enabled = False Then Exit Sub
        
    
    Select Case Button.Key
        
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call LoadRecord
            
            
    End Select
End Sub


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property


Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property


Public Property Let UpdFlg(InUpdFlg As Boolean)
    wbUpdFlg = InUpdFlg
End Property


