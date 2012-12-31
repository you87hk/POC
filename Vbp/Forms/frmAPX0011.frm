VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmAPX0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPX0011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10061.66
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9735
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
      Height          =   5655
      Left            =   120
      OleObjectBlob   =   "frmAPX0011.frx":0442
      TabIndex        =   0
      Top             =   1320
      Width           =   9855
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
Attribute VB_Name = "frmAPX0011"
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


Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String





Private Const SDOCLINE = 0
Private Const SBOOKCODE = 1
Private Const SWHSCODE = 2
Private Const SLOTNO = 3
Private Const SQTY = 4
Private Const SRSVQTY = 5
Private Const SSCHQTY = 6
Private Const SDISPER = 7
Private Const SAMT = 8
Private Const SNET = 9
Private Const SDUMMY = 10


Public Property Let InDocID(inDoc As Long)
     wlDocID = inDoc
End Property


Public Property Let InCusNo(inDocCus As String)
     wsCusNo = inDocCus
End Property




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF5
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
    
    waResult.ReDim 0, -1, SDOCLINE, SNET
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
    
    
    lblDspDocNo.Caption = Get_TableInfo("SOASNHD", "SNHDDOCID = " & wlDocID, "SNHDDOCNO")
    lblDspCusNo.Caption = wsCusNo & " - " & Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(wsCusNo) & "'", "CUSNAME")
    
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "APX0011"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblCusNo.Caption = Get_Caption(waScrItm, "CUSNO")
        
    
    With tblDetail
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SBOOKCODE).Caption = Get_Caption(waScrItm, "SBOOKCODE")
        .Columns(SWHSCODE).Caption = Get_Caption(waScrItm, "SWHSCODE")
        .Columns(SLOTNO).Caption = Get_Caption(waScrItm, "SLOTNO")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SRSVQTY).Caption = Get_Caption(waScrItm, "SRSVQTY")
        .Columns(SSCHQTY).Caption = Get_Caption(waScrItm, "SSCHQTY")
        .Columns(SDISPER).Caption = Get_Caption(waScrItm, "SDISPER")
        .Columns(SAMT).Caption = Get_Caption(waScrItm, "SAMT")
        .Columns(SNET).Caption = Get_Caption(waScrItm, "SNET")
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
                Case SNET
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
                Case SNET
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






Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

     wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case SBOOKCODE
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = Get_TableInfo("MSTITEM", "ITMCODE = '" & Set_Quote(.Columns(SBOOKCODE).Text) & "'", IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME"))
    
                Case SWHSCODE
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = Get_TableInfo("MSTWAREHOUSE", "WHSCODE = '" & Set_Quote(.Columns(SWHSCODE).Text) & "'", "WHSDESC")
                    
                 
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub


Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SDOCLINE To SDUMMY
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SDOCLINE
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Width = 1000
                Case SBOOKCODE
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Width = 1500
                Case SWHSCODE
                   .Columns(wiCtr).Width = 1000
                   .Columns(wiCtr).DataWidth = 10
                Case SLOTNO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 20
                Case SRSVQTY
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                Case SSCHQTY
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                Case SQTY
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case SDISPER
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SAMT
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SNET
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wdQty As Double
    Dim wdPrice As Double
    Dim wdAmt As Double
    Dim wdDis As Double
    Dim wdNet As Double
    
    LoadRecord = False
    Me.MousePointer = vbHourglass
    
    
        wsSql = "SELECT SNDTDOCLINE, ITMCODE, SNDTWHSCODE, SNDTLOTNO, SNDTQTY, SNDTRSVQTY, SNDTSCHQTY, "
        wsSql = wsSql & " SNDTDISPER, SNDTAMT, SNDTNET "
        wsSql = wsSql & " FROM  soaSNDT, mstITEM "
        wsSql = wsSql & " WHERE SNDTDOCID = " & wlDocID
        wsSql = wsSql & " AND SNDTITEMID = ITMID "
        wsSql = wsSql & " ORDER BY SNDTDOCLINE "
    
     rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    
    
    With waResult
    .ReDim 0, -1, SDOCLINE, SNET
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

     .AppendRows
        waResult(.UpperBound(1), SDOCLINE) = Format(ReadRs(rsRcd, "SNDTDOCLINE"), "000")
        waResult(.UpperBound(1), SBOOKCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SWHSCODE) = ReadRs(rsRcd, "SNDTWHSCODE")
        waResult(.UpperBound(1), SLOTNO) = ReadRs(rsRcd, "SNDTLOTNO")
        waResult(.UpperBound(1), SRSVQTY) = Format(To_Value(ReadRs(rsRcd, "SNDTRSVQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SSCHQTY) = Format(To_Value(ReadRs(rsRcd, "SNDTSCHQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "SNDTQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SDISPER) = Format(To_Value(ReadRs(rsRcd, "SNDTDISPER")), gsAmtFmt)
        waResult(.UpperBound(1), SAMT) = Format(To_Value(ReadRs(rsRcd, "SNDTAMT")), gsAmtFmt)
        waResult(.UpperBound(1), SNET) = Format(To_Value(ReadRs(rsRcd, "SNDTNET")), gsAmtFmt)
        
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




