VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmEIM0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmEIM0011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10061.66
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnPrint 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin VB.Label lblDspDocNo 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   5535
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
      Height          =   6015
      Left            =   120
      OleObjectBlob   =   "frmEIM0011.frx":0442
      TabIndex        =   0
      Top             =   960
      Width           =   9855
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   9855
   End
End
Attribute VB_Name = "frmEIM0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private wbErr As Boolean

Private wsTrnCd As String
Private wsLocCode As String
Private wsDteTim As String
Private wsUsrId As String
Private wiExit As Boolean
Private wsStatus As String

Private wsFormCaption As String
Private wsFormID As String


Private Const SDOCLINE = 0
Private Const SBOOKCODE = 1
Private Const SBOOKNAME = 2
Private Const SWHSCODE = 3
Private Const SQTY = 4
Private Const STRNCODE = 5
Private Const SDUMMY = 6

Public Property Let InTrnCd(inTrn As String)
     wsTrnCd = inTrn
End Property

Public Property Let InLocCode(inLoc As String)
     wsLocCode = inLoc
End Property

Public Property Let InUsrID(inUsr As String)
     wsUsrId = inUsr
End Property

Public Property Let inDteTim(inDT As String)
     wsDteTim = inDT
End Property

Public Property Let InStatus(inSts As String)
     wsStatus = inSts
End Property


Private Sub btnPrint_Click()
Call cmdPrint
End Sub

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
    
    waResult.ReDim 0, -1, SDOCLINE, STRNCODE
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
    
    
    lblDspDocNo.Caption = wsLocCode & " - " & Get_TableInfo("MSTLOCATION", "LOCCODE = '" & wsLocCode & "'", "LOCNAME")
    
    
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    Set frmEIM0011 = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "EIM0011"
    
    If wsTrnCd = "EM" Then
        btnPrint.Visible = True
    Else
        btnPrint.Visible = False
    End If
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    btnPrint.Caption = Get_Caption(waScrItm, "PRINT")
 
    With tblDetail
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SBOOKCODE).Caption = Get_Caption(waScrItm, "SBOOKCODE")
        .Columns(SBOOKNAME).Caption = Get_Caption(waScrItm, "SBOOKNAME")
        
        .Columns(SWHSCODE).Caption = Get_Caption(waScrItm, "SWHSCODE")
        .Columns(STRNCODE).Caption = Get_Caption(waScrItm, "STRNCODE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        
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
                Case STRNCODE
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
                Case STRNCODE
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
                    lblDspItmDesc.Caption = .Columns(SBOOKNAME).Text
                Case SBOOKNAME
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SBOOKNAME).Text
    
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
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 2000
                Case SBOOKCODE
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Width = 1500
                Case SBOOKNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 4500
                Case SWHSCODE
                   .Columns(wiCtr).Width = 1000
                   .Columns(wiCtr).DataWidth = 10
                   .Columns(wiCtr).Visible = False
                Case SQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case STRNCODE
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Visible = False
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
    
        wsSql = "SELECT EIHDDOCNO, ITMCODE, EIDTITEMDESC, EIDTITEMID, EIDTWHSCODE, EIDTLOTNO, EIDTQTY, CASE EIHDTRNCODE WHEN 'EM' THEN 'Y' ELSE 'N' END TRNCODE "
        wsSql = wsSql & " FROM  INEIHD, INEIDT, MSTITEM "
        wsSql = wsSql & " WHERE EIDTDOCID = EIHDDOCID "
        wsSql = wsSql & " AND EIDTITEMID = ITMID "
        wsSql = wsSql & " AND EIHDLOCCODE = '" & Set_Quote(wsLocCode) & "' "
        wsSql = wsSql & " AND EIHDUSRID = '" & Set_Quote(wsUsrId) & "' "
        wsSql = wsSql & " AND EIHDDTETIM = '" & Change_SQLDate(wsDteTim) & "' "
        wsSql = wsSql & " AND EIHDSTATUS = '" & wsStatus & "' "
        If wsTrnCd = "EM" Then
        wsSql = wsSql & " AND EIHDTRNCODE = 'EM' "
        Else
        wsSql = wsSql & " AND EIHDTRNCODE <> 'EM' "
        End If
        
        wsSql = wsSql & " ORDER BY EIHDDOCNO, EIDTDOCLINE "
     
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    
    
    With waResult
    .ReDim 0, -1, SDOCLINE, STRNCODE
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
     .AppendRows
        waResult(.UpperBound(1), SDOCLINE) = ReadRs(rsRcd, "EIHDDOCNO")
        waResult(.UpperBound(1), SBOOKCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SBOOKNAME) = ReadRs(rsRcd, "EIDTITEMDESC")
        waResult(.UpperBound(1), SWHSCODE) = ReadRs(rsRcd, "EIDTWHSCODE")
        waResult(.UpperBound(1), STRNCODE) = ReadRs(rsRcd, "TRNCODE")
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "EIDTQTY")), gsQtyFmt)
        
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



Private Function InputValidation() As Boolean
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If To_Value(waResult(wlCtr, SQTY)) <> 0 Then
                wiEmptyGrid = False
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有差異資料!"
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


Private Sub cmdPrint()
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsRptTitle As String
    
    
    Me.MousePointer = vbHourglass
    
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
   wsRptTitle = "書展預訂單"
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    'Create Stored Procedure String
    
    wsSql = "EXEC usp_RPTEIM002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wsRptTitle & "', "
    wsSql = wsSql & gsLangID
    
   
    wsRptName = "CRPTEIM002"
    
    
    NewfrmPrint.ReportID = "EIM002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "EIM002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    



    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
    

    
End Sub
