VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmAPR0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPR0011.frx":0000
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
      Begin VB.Label lblReceived 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblDspReceived 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDspCusNo 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label lblDspDocNo 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblCusNo 
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
      OleObjectBlob   =   "frmAPR0011.frx":0442
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
            Picture         =   "frmAPR0011.frx":8B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":944F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":9D29
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":A17B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":A5CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":A8E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":AD39
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":B18B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":B4A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":B7BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":BC11
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":C4ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":C815
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":CC69
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":CF85
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":D2A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":D6F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":DA11
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":DD31
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":E051
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR0011.frx":E4A5
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
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pick"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CalRsv"
            ImageIndex      =   14
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
Attribute VB_Name = "frmAPR0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private wbErr As Boolean

Private wsCusNo As String
Private wlDocID As Long
Private wlLastRow As Long

Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String
Private wbUpdFlg As Boolean
Private wsTrnCd As String


Private Const tcPick = "Pick"
Private Const tcCalRsv = "CalRsv"
Private Const tcRefresh = "Refresh"
Private Const tcExit = "Exit"

Private Const SSEL = 0
Private Const SDOCLINE = 1
Private Const SITMCODE = 2
Private Const SITMNAME = 3
Private Const SITMTYPE = 4
Private Const SUPRICE = 5
Private Const SQTY = 6
Private Const SSCHQTY = 7
Private Const SSOH = 8
Private Const SRSVQTY = 9
Private Const SDUMMY = 10
Private Const SID = 11



Public Property Let InDocID(inDoc As Long)
     wlDocID = inDoc
End Property


Public Property Let InCusNo(inDocCus As String)
wsCusNo = inDocCus
End Property


Private Sub cmdReserve()
    
  
    Call LoadRecord
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF7
            Call LoadRecord
        
        Case vbKeyF2
        If tbrProcess.Buttons(tcCalRsv).Enabled = False Then Exit Sub
            Call cmdReserve
            
        Case vbKeyF3
        If tbrProcess.Buttons(tcPick).Enabled = False Then Exit Sub
            Call cmdPick
            
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    
    
  MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Grid
    Set_tbrProcess
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
    
    waResult.ReDim 0, -1, SSEL, SID
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
    
    If wbUpdFlg = False Then tbrProcess.Buttons(tcCalRsv).Enabled = False
    
    Select Case wsTrnCd
    Case "SN"
    lblDspDocNo.Caption = Get_TableInfo("SOASNHD", "SNHDDOCID = " & wlDocID, "SNHDDOCNO")
    Case "SO"
    lblDspDocNo.Caption = Get_TableInfo("SOASOHD", "SOHDDOCID = " & wlDocID, "SOHDDOCNO")
    Case "SP"
    lblDspDocNo.Caption = Get_TableInfo("SOASPHD", "SPHDDOCID = " & wlDocID, "SPHDDOCNO")
    Case "SW"
    lblDspDocNo.Caption = Get_TableInfo("SOASWHD", "SWHDDOCID = " & wlDocID, "SWHDDOCNO")
    lblDspReceived.Caption = Get_TableInfo("SOASWDT, MSTSALESMAN", "SWDTWORKERID = SALEID AND SWDTDOCID = " & wlDocID, "SALENAME")
    End Select
    
    lblDspCusNo.Caption = wsCusNo & " - " & Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(wsCusNo) & "'", "CUSNAME")
    
  '  wbUpdate = False
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    Set frmAPR0011 = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
   If wsFormID = "APR0061" Then
   lblReceived.Visible = True
   lblDspReceived.Visible = True
   Else
   lblReceived.Visible = False
   lblDspReceived.Visible = False
   End If
   
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblCusNo.Caption = Get_Caption(waScrItm, "CUSNO")
    
    lblReceived.Caption = Get_Caption(waScrItm, "RECEIVED")
        
        
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SSOH).Caption = Get_Caption(waScrItm, "SSOH")
        .Columns(SRSVQTY).Caption = Get_Caption(waScrItm, "SRSVQTY")
        .Columns(SSCHQTY).Caption = Get_Caption(waScrItm, "SSCHQTY")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
    End With
    
    tbrProcess.Buttons(tcPick).ToolTipText = Get_Caption(waScrItm, tcPick) & "(F3)"
    tbrProcess.Buttons(tcCalRsv).ToolTipText = Get_Caption(waScrItm, tcCalRsv) & "(F2)"
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
    
  
       
    With tblDetail
        Select Case ColIndex
            
           
                
           Case SRSVQTY
                
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
               
               If To_Value(.Columns(ColIndex).Text) > 0 Then
               .Columns(SSEL) = "-1"
               End If
               
           End Select
            
           '  If .Columns(ColIndex).Text <> OldValue Then
            '    wbUpdate = True
            ' End If
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



Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
       
        Case vbKeyReturn
            Select Case .Col
                Case SRSVQTY
                    KeyCode = vbKeyDown
                    .Col = SRSVQTY
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SSEL Then
                   .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SRSVQTY
                    KeyCode = vbKeyDown
                    .Col = SSEL
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
        
        Case SRSVQTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
       
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
               
                Case SQTY
                    
                    Call Chk_grdQty(.Columns(SQTY).Text)
                    
                   
            End Select
            lblDspItmDesc.Caption = .Columns(.Col).Text
        End If
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

 '   If To_Value(inCode) = 0 Then
 '       gsMsg = "數量必需大於零!"
 '       MsgBox gsMsg, vbOKOnly, gsTitle
 '       Chk_grdQty = False
 '       Exit Function
 '   End If
    
    
    
    
End Function

Private Function Chk_grdRsvQty(inCode As String) As Boolean
    
    Chk_grdRsvQty = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdRsvQty = False
        Exit Function
    End If

    If To_Value(inCode) = 0 Then
        gsMsg = "數量必需大於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdRsvQty = False
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
        
        For wiCtr = SSEL To SID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SSEL
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).Locked = False
                Case SDOCLINE
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Width = 500
                Case SITMCODE
                    .Columns(wiCtr).DataWidth = 30
                    .Columns(wiCtr).Width = 2000
                Case SITMNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 1500
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
                Case SRSVQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = False
                Case SSCHQTY
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
    
    Case "SN"
    
        wsSQL = "SELECT SNDTID SID, SNDTDOCLINE DOCLINE, SNDTITEMID ITEMID, ITMCODE, SNDTITEMDESC ITEMDESC, ITMITMTYPECODE, SNDTUPRICE UPRICE, SNDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " SNDTTOTQTY QTY, SNDTRSVQTY RSVQTY, SNDTSCHQTY SCHQTY "
        wsSQL = wsSQL & " FROM soaSNDT, mstITEM "
        wsSQL = wsSQL & " WHERE SNDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND SNDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY SNDTDOCLINE "
    
    Case "SO"
    
        wsSQL = "SELECT SODTITEMID SID, SODTITEMID ITEMID, ITMCODE, SODTITEMDESC ITEMDESC, ITMITMTYPECODE, SODTUPRICE UPRICE, 'P' LNTYPE, "
        wsSQL = wsSQL & " SUM(SODTTOTQTY) QTY, SUM(SODTRSVQTY) RSVQTY, SUM(SODTSCHQTY) SCHQTY "
        wsSQL = wsSQL & " FROM soaSODT, mstITEM "
        wsSQL = wsSQL & " WHERE SODTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND SODTITEMID = ITMID "
        wsSQL = wsSQL & " GROUP BY SODTITEMID,  SODTITEMID, ITMCODE, SODTITEMDESC, ITMITMTYPECODE, SODTUPRICE "
    
    Case "SP"
    
        wsSQL = "SELECT SPDTID SID, SPDTDOCLINE DOCLINE, SPDTITEMID ITEMID, ITMCODE, SPDTITEMDESC ITEMDESC, ITMITMTYPECODE, SPDTUPRICE UPRICE, SPDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " SPDTQTY QTY, SPDTBALQTY RSVQTY, SPDTOUTQTY SCHQTY "
        wsSQL = wsSQL & " FROM soaSPDT, mstITEM "
        wsSQL = wsSQL & " WHERE SPDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND SPDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY SPDTDOCLINE "
    
    Case "SW"
    
        wsSQL = "SELECT SWDTID SID, SWDTDOCLINE DOCLINE, SWDTITEMID ITEMID, ITMCODE, SWDTITEMDESC ITEMDESC, ITMITMTYPECODE, SWDTUPRICE UPRICE, SWDTLNTYPE LNTYPE, "
        wsSQL = wsSQL & " SWDTQTY QTY, SWDTBALQTY RSVQTY, SWDTOUTQTY SCHQTY "
        wsSQL = wsSQL & " FROM soaSWDT, mstITEM "
        wsSQL = wsSQL & " WHERE SWDTDOCID = " & wlDocID
        wsSQL = wsSQL & " AND SWDTITEMID = ITMID "
        wsSQL = wsSQL & " ORDER BY SWDTDOCLINE "
    
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
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

    If ReadRs(rsRcd, "LNTYPE") = "T" And wsTrnCd = "SN" Then
    wdQty = 0
    wsItmCd = "*" & ReadRs(rsRcd, "ITMCODE")
    Else
    wdQty = Get_StockAvailable(To_Value(ReadRs(rsRcd, "ITEMID")), "", "")
    wsItmCd = ReadRs(rsRcd, "ITMCODE")
    End If
    
     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCLINE) = Format(wiCtr, "000")
        waResult(.UpperBound(1), SITMCODE) = wsItmCd
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "ITEMDESC")
        waResult(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
        waResult(.UpperBound(1), SUPRICE) = Format(To_Value(ReadRs(rsRcd, "UPRICE")), gsUprFmt)
        waResult(.UpperBound(1), SRSVQTY) = Format(To_Value(ReadRs(rsRcd, "RSVQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SSCHQTY) = Format(To_Value(ReadRs(rsRcd, "SCHQTY")), gsQtyFmt)
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
        
        Case tcPick
           Call cmdPick
        
        Case tcCalRsv
            Call cmdReserve
           
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call LoadRecord
            
            
    End Select
End Sub

Private Sub Set_tbrProcess()

With tbrProcess
    
    Select Case wsFormID
    Case "APR0011"
    
    .Buttons(tcCalRsv).Enabled = True
    .Buttons(tcPick).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    Case "APR0021"
    
    .Buttons(tcCalRsv).Enabled = True
    .Buttons(tcPick).Enabled = True
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    Case "APR0031", "APR0061"
    
    .Buttons(tcCalRsv).Enabled = False
    .Buttons(tcPick).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    End Select
    
    If wbUpdFlg = False Then
    .Buttons(tcCalRsv).Enabled = False
    .Buttons(tcPick).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcExit).Enabled = True
    End If
    
End With

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

Private Function InputValidation() As Boolean
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    wiEmptyGrid = True
    wlLastRow = 0
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SSEL)) = "-1" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
                wlLastRow = wlLastRow + 1
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
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
        
        If Chk_grdRsvQty(waResult(LastRow, SRSVQTY)) = False Then
                .Col = SRSVQTY
               Exit Function
        End If
    
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function


Private Sub cmdPick()

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlLineNo As Long
    Dim wlHDID As Long
     
    On Error GoTo cmdPick_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    gsMsg = "你是否確認要轉換配貨單?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

 
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    wlLineNo = 1
    wlHDID = 0
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_CRTPICKLST"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wlDocID)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 3, wlHDID)
                Call SetSPPara(adcmdSave, 4, wlLineNo)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, SITMNAME))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, SUPRICE))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, SRSVQTY))
                Call SetSPPara(adcmdSave, 8, wsFormID)
                Call SetSPPara(adcmdSave, 9, gsUserID)
                Call SetSPPara(adcmdSave, 10, wsGenDte)
                Call SetSPPara(adcmdSave, 11, IIf(wlLastRow = wlLineNo, "Y", "N"))
                
                adcmdSave.Execute
                wlHDID = GetSPPara(adcmdSave, 12)
                wsDocNo = GetSPPara(adcmdSave, 13)
                wlLineNo = wlLineNo + 1
            End If
        Next
    End If
    
     
    cnCon.CommitTrans
    
  
    gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
        

    
    Set adcmdSave = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    Exit Sub
    
cmdPick_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub
