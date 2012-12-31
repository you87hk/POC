VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmSOLOG 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmSOLOG.frx":0000
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
      OleObjectBlob   =   "frmSOLOG.frx":0442
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
            Picture         =   "frmSOLOG.frx":7E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":874B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":9025
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":9477
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":98C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":9BE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":A035
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":A487
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":A7A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":AABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":AF0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":B7E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":BB11
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":BF65
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":C281
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":C59D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":C9F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":CD0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":D02D
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":D34D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOLOG.frx":D7A1
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
Attribute VB_Name = "frmSOLOG"
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

Private wsTrnCd As String


Private Const tcRefresh = "Refresh"
Private Const tcExit = "Exit"

Private Const SDTETIM = 0
Private Const SUSRID = 1
Private Const sType = 2
Private Const SACT = 3
Private Const SDOCNO = 4
Private Const SKEY = 5
Private Const SFIELD = 6
Private Const SOLD = 7
Private Const SNEW = 8



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
    
    waResult.ReDim 0, -1, SDTETIM, SNEW
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
    Case "SO"
    lblDspDocNo.Caption = Get_TableInfo("SOASOHD", "SOHDDOCID = " & wlDocID, "SOHDDOCNO")
    End Select
    
    lblDspCusNo.Caption = wsCusNo & " - " & Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(wsCusNo) & "'", "CUSNAME")
    
  '  wbUpdate = False
    
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    
    Set waScrItm = Nothing
    Set waResult = Nothing
    Set frmSOLOG = Nothing
    
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
    lblCusNo.Caption = Get_Caption(waScrItm, "CUSNO")
    
    With tblDetail
        .Columns(SDTETIM).Caption = Get_Caption(waScrItm, "SDTETIM")
        .Columns(SUSRID).Caption = Get_Caption(waScrItm, "SUSRID")
        .Columns(sType).Caption = Get_Caption(waScrItm, "STYPE")
        .Columns(SACT).Caption = Get_Caption(waScrItm, "SACT")
        .Columns(SKEY).Caption = Get_Caption(waScrItm, "SKEY")
        .Columns(SFIELD).Caption = Get_Caption(waScrItm, "SFIELD")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SOLD).Caption = Get_Caption(waScrItm, "SOLD")
        .Columns(SNEW).Caption = Get_Caption(waScrItm, "SNEW")
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
                Case SNEW
                    KeyCode = vbKeyDown
                    .Col = SDTETIM
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SDTETIM Then
                   .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SNEW
                    KeyCode = vbKeyDown
                    .Col = SDTETIM
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
               
                Case SKEY
                    
                    
                   
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
        .MultipleLines = 0
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SDTETIM To SNEW
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SDTETIM
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Width = 2300
                Case SUSRID
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1000
                Case sType
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1000
                Case SACT
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 800
                Case SDOCNO
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                Case SKEY
                    .Columns(wiCtr).Width = 1800
                    .Columns(wiCtr).DataWidth = 30
                Case SFIELD
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                Case SOLD
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                Case SNEW
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 15
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim wiCtr As Long
    
    LoadRecord = False
    Me.MousePointer = vbHourglass
    
    Set cmd.ActiveConnection = cnCon
        
    cmd.CommandText = "USP_LOADSOLOG"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    Call SetSPPara(cmd, 1, wlDocID)
    Call SetSPPara(cmd, 2, "")
    
    
    
    rsRcd.CursorLocation = adUseClient
    rsRcd.Open cmd, , adOpenStatic, adLockReadOnly

    ' rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    wiCtr = 1
    With waResult
    .ReDim 0, -1, SDTETIM, SNEW
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

    
     .AppendRows
        waResult(.UpperBound(1), SDTETIM) = ReadRs(rsRcd, "TMPDATETIM")
        waResult(.UpperBound(1), SUSRID) = ReadRs(rsRcd, "TMPUSRID")
        waResult(.UpperBound(1), sType) = ReadRs(rsRcd, "TMPTYPE")
        waResult(.UpperBound(1), SACT) = ReadRs(rsRcd, "TMPACT")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "TMPDOCNO")
        waResult(.UpperBound(1), SFIELD) = ReadRs(rsRcd, "TMPFIELD")
        waResult(.UpperBound(1), SKEY) = ReadRs(rsRcd, "TMPKEY")
        waResult(.UpperBound(1), SOLD) = ReadRs(rsRcd, "TMPOLD")
        waResult(.UpperBound(1), SNEW) = ReadRs(rsRcd, "TMPNEW")
        wiCtr = wiCtr + 1
        
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    rsRcd.Close
    
    
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
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

