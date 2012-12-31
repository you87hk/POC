VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmHHIM0011 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmHHIM0011.frx":0000
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
      Begin VB.Label lblDspDteTim 
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
      Begin VB.Label lblDteTim 
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
      OleObjectBlob   =   "frmHHIM0011.frx":0442
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
            Picture         =   "frmHHIM0011.frx":8F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":97DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":A0B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":A50B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":A95D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":AC77
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":B0C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":B51B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":B835
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":BB4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":BFA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":C87D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":CBA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":CFF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":D315
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":D631
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":DA85
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":DDA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":E0C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":E3E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM0011.frx":E835
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
Attribute VB_Name = "frmHHIM0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private wbErr As Boolean


Private wsDocID As String
Private wsDteTim As String
Private wlLastRow As Long

Private wiExit As Boolean

Private wsFormCaption As String
Private wsFormID As String

Private wsTrnCd As String


Private Const tcRefresh = "Refresh"
Private Const tcExit = "Exit"

Private Const SSEL = 0
Private Const SDOCNO = 1
Private Const sType = 2
Private Const SLOCCODE = 3
Private Const SITMCODE = 4
Private Const SHHQTY = 5
Private Const SIMQTY = 6
Private Const SSTAFF = 7
Private Const SLINE = 8
Private Const SABC = 9
Private Const SDUMMY = 10
Private Const SID = 11



Public Property Let InDocID(inDoc As String)
     wsDocID = inDoc
End Property

Public Property Let inDteTim(inDte As String)
     wsDteTim = inDte
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
    
    
    lblDspDocNo.Caption = wsDocID
    lblDspDteTim.Caption = wsDteTim
    
  '  wbUpdate = False
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    Set frmHHIM0011 = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblDteTim.Caption = Get_Caption(waScrItm, "DTETIM")
    

    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(sType).Caption = Get_Caption(waScrItm, "STYPE")
        .Columns(SLOCCODE).Caption = Get_Caption(waScrItm, "SLOCCODE")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SHHQTY).Caption = Get_Caption(waScrItm, "SHHQTY")
        .Columns(SIMQTY).Caption = Get_Caption(waScrItm, "SIMQTY")
        .Columns(SSTAFF).Caption = Get_Caption(waScrItm, "SSTAFF")
        .Columns(SLINE).Caption = Get_Caption(waScrItm, "SLINE")
        .Columns(SABC).Caption = Get_Caption(waScrItm, "SABC")
        
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
                Case SABC
                    KeyCode = vbKeyDown
                    .Col = SSEL
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
                Case SABC
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
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).Visible = False
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 1200
                Case sType
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Width = 800
                Case SLOCCODE
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1000
                Case SITMCODE
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 30
                Case SHHQTY
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SIMQTY
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SSTAFF
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SLINE
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                Case SABC
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
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
    
    LoadRecord = False
    Me.MousePointer = vbHourglass
    

    
    
    wsSQL = "SELECT HHDOCNO DOCNO, HHLOCCODE LOCCODE, HHITMCODE ITMCODE, HHTYPE TYPE, HHQTY, "
    wsSQL = wsSQL & " HHIMQTY IMQTY,HHSTAFF, HHLINE, HHABCFLG "
    wsSQL = wsSQL & " FROM syshhim001 "
    wsSQL = wsSQL & " WHERE HHNO = '" & wsDocID & "' "
    wsSQL = wsSQL & " ORDER BY HHITMCODE "
   
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

    
     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), sType) = ReadRs(rsRcd, "TYPE")
        waResult(.UpperBound(1), SLOCCODE) = ReadRs(rsRcd, "LOCCODE")
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SHHQTY) = Format(To_Value(ReadRs(rsRcd, "HHQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SIMQTY) = Format(To_Value(ReadRs(rsRcd, "IMQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SSTAFF) = ReadRs(rsRcd, "HHSTAFF")
        waResult(.UpperBound(1), SLINE) = Format(To_Value(ReadRs(rsRcd, "HHLINE")), gsQtyFmt)
        waResult(.UpperBound(1), SABC) = Format(To_Value(ReadRs(rsRcd, "HHABCFLG")), gsQtyFmt)
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DOCNO")
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

Private Sub Set_tbrProcess()

With tbrProcess
    
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
End With

End Sub

Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property


Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property


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
        
    
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function

