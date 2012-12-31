VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmB0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "B0011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   6742.814
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      Begin VB.Label lblDspBookName 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label lblDspISBN 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblBookName 
         Caption         =   "BOOKNAME"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblISBN 
         Caption         =   "ISBN"
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
      OleObjectBlob   =   "B0011.frx":0442
      TabIndex        =   0
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   6495
   End
End
Attribute VB_Name = "frmB0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private wbErr As Boolean

Private wlItemID As Long
Private wsISBN As String
Private wsBookName As String

Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String

Private Const DOCDATE = 0
Private Const DOCNO = 1
Private Const OLDPRICE = 2
Private Const NEWPRICE = 3
Private Const DISC = 4
Private Const SDUMMY = 5

Public Property Let InItemID(inID As Long)
     wlItemID = inID
End Property

Public Property Let inISBN(inBookISBN As String)
     wsISBN = inBookISBN
End Property

Public Property Let InBookName(inBkName As String)
     wsBookName = inBkName
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
    
    waResult.ReDim 0, -1, DOCDATE, DISC
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
    
    'lblDspDocNo.Caption = Get_TableInfo("SOASNHD", "SNHDDOCID = " & wlDocID, "SNHDDOCNO")
    'lblDspCusNo.Caption = wsCusNo & " - " & Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(wsCusNo) & "'", "CUSNAME")
    lblDspISBN = wsISBN
    lblDspBookName = wsBookName
    
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
    
    wsFormID = "B0011"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblISBN.Caption = Get_Caption(waScrItm, "ISBN")
    lblBookName.Caption = Get_Caption(waScrItm, "BOOKNAME")
        
    
    With tblDetail
        .Columns(DOCDATE).Caption = Get_Caption(waScrItm, "DOCDATE")
        .Columns(DOCNO).Caption = Get_Caption(waScrItm, "DOCNO")
        .Columns(OLDPRICE).Caption = Get_Caption(waScrItm, "OLDPRICE")
        .Columns(NEWPRICE).Caption = Get_Caption(waScrItm, "NEWPRICE")
        .Columns(DISC).Caption = Get_Caption(waScrItm, "DISC")
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
                Case DISC
                    KeyCode = vbKeyDown
                    .Col = DOCDATE
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> DOCDATE Then
                   .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case DISC
                    KeyCode = vbKeyDown
                    .Col = DOCDATE
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
                'Case DISC
                '    lblDspItmDesc.Caption = ""
                '    lblDspItmDesc.Caption = Get_TableInfo("MSTWAREHOUSE", "WHSCODE = '" & Set_Quote(.Columns(SWHSCODE).Text) & "'", "WHSDESC")
                    
                 
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
        
        For wiCtr = DOCDATE To SDUMMY
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case DOCDATE
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1300
                Case DOCNO
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Width = 1500
                Case OLDPRICE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case NEWPRICE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case DISC
                    .Columns(wiCtr).Width = 1100
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case SDUMMY
                    .Columns(wiCtr).Width = 10
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
    
    
    wsSql = "SELECT PRICHGDOCNO, PRICHGDOCDATE, PRICHGDISPERIN, PRICHGDOCTYPE, "
    wsSql = wsSql & " PRICHGDTDEFAULTPRICE, PRICHGDTUNITPRICE "
    wsSql = wsSql & " FROM  MstPriceChange, MstPriceChangeDT "
    wsSql = wsSql & " WHERE PRICHGDTITEMID = " & CStr(wlItemID)
    wsSql = wsSql & " AND PRICHGDOCID = PRICHGDTDOCID "
    wsSql = wsSql & " ORDER BY PRICHGDOCDATE "
    
     rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    With waResult
    .ReDim 0, -1, DOCDATE, DISC
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

     .AppendRows
        waResult(.UpperBound(1), DOCDATE) = ReadRs(rsRcd, "PRICHGDOCDATE")
        waResult(.UpperBound(1), DOCNO) = ReadRs(rsRcd, "PRICHGDOCNO")
        waResult(.UpperBound(1), OLDPRICE) = ReadRs(rsRcd, "PRICHGDTDEFAULTPRICE")
        waResult(.UpperBound(1), NEWPRICE) = ReadRs(rsRcd, "PRICHGDTUNITPRICE")
        waResult(.UpperBound(1), DISC) = IIf(ReadRs(rsRcd, "PRICHGDOCTYPE") = "A", "+", "-") & Format(ReadRs(rsRcd, "PRICHGDISPERIN"), gsAmtFmt)
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




