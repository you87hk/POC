VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEIM001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmEIM001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSelect 
      Height          =   525
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Width           =   3975
      Begin VB.OptionButton optDocType 
         Caption         =   "SO"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   200
         Width           =   1335
      End
      Begin VB.OptionButton optDocType 
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   200
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fmeFilePath 
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   7815
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         Width           =   288
      End
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Left            =   1440
         TabIndex        =   5
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   240
         Width           =   5820
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Export Sheet Path:"
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   240
         Width           =   1110
      End
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10200
      OleObjectBlob   =   "frmEIM001.frx":0442
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   4575
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6855
      Left            =   0
      OleObjectBlob   =   "frmEIM001.frx":2B45
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   11400
      Top             =   360
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
            Picture         =   "frmEIM001.frx":A854
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":B12E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":BA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":BE5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":C2AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":C5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":CA18
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":CE6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":D184
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":D49E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":D8F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":E1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":E4F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":E948
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":EC64
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":EF80
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":F3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":F6F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":FA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":FE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEIM001.frx":1017C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OK"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Can"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Import"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog cdFilePath 
      Left            =   11520
      Top             =   1200
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmEIM001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Private wbErr As Boolean
Private wsExcelPath As String
Private wsDteTim As String


Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wsTrnCd As String
Private wiActRow As Integer


Private Const tcOK = "OK"
Private Const tcCan = "Can"
Private Const tcImport = "Import"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const SUSRID = 1
Private Const SDTETIM = 2
Private Const SLOCCODE = 3
Private Const SLOCNAME = 4
Private Const SQTY = 5
Private Const STRNCODE = 6
Private Const SDUMMY = 7
Private Const SID = 8


Private Sub cmdFilePath_Click()

      On Error Resume Next

    If Trim(txtFilePath.Text) <> "" Then
        If Dir(txtFilePath.Text) <> "" Then
            'Dialog.dirDirectory.Path = txtFilePath.Text
            cdFilePath.InitDir = wsExcelPath
        End If
    End If
    
    With cdFilePath
    .DialogTitle = "Open A text File"
    .Filter = "Text File (*.TXT)|*.TXT"
    .CancelError = True
    .FileName = vbNullString
    .ShowOpen
    If Err.Number <> cdlCancel Then
        txtFilePath.Text = .FileName
    End If
    End With
   
   On Error GoTo 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6
           Call cmdSave(1)
           
        Case vbKeyF7
            Call cmdSave(2)
             
          
        Case vbKeyF8
           Call cmdImport
           
        Case vbKeyF3
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF5
            Call LoadRecord
        
        Case vbKeyF9
           Call cmdSelect(1)
           
        Case vbKeyF10
           Call cmdSelect(0)
        
      
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

Private Sub cmdImport()
 
    
     
  If chk_txtFilePath = False Then Exit Sub
     
  MousePointer = vbHourglass
  
  wsDteTim = Change_SQLDate(Now)
  Call ImportEONew(gsUserID, wsDteTim, txtFilePath.Text)
 'Call ImportEO(gsUserID, wsDteTim, txtFilePath.Text)
  gsMsg = "匯入完成!"
  MsgBox gsMsg, vbOKOnly, gsTitle
  
  Call LoadRecord
    
    
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
    
    tblCommon.Visible = False
    wiExit = False
    wsTrnCd = ""
      
    If InStr(gsComPath, ":\") Or InStr(gsComPath, "\\") Then
        wsExcelPath = gsComPath
    Else
        wsExcelPath = App.Path & "\" & gsComPath
    End If
    
    txtFilePath.Text = wsExcelPath & "IMPORT.TXT"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmEIM001 = Nothing


    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    optDocType(0).Value = True
    
  '  wsFormID = "EIM001"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblFilePath.Caption = Get_Caption(waScrItm, "FILEPATH")
    optDocType(0).Caption = Get_Caption(waScrItm, "OPT1")
    optDocType(1).Caption = Get_Caption(waScrItm, "OPT2")
        
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SLOCNAME).Caption = Get_Caption(waScrItm, "SLOCNAME")
        .Columns(SLOCCODE).Caption = Get_Caption(waScrItm, "SLOCCODE")
        .Columns(SUSRID).Caption = Get_Caption(waScrItm, "SUSRID")
        .Columns(SDTETIM).Caption = Get_Caption(waScrItm, "SDTETIM")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(STRNCODE).Caption = Get_Caption(waScrItm, "STRNCODE")

    End With
    
    
    tbrProcess.Buttons(tcOK).ToolTipText = Get_Caption(waScrToolTip, tcOK) & "(F6)"
    tbrProcess.Buttons(tcImport).ToolTipText = Get_Caption(waScrToolTip, tcImport) & "(F8)"
    tbrProcess.Buttons(tcCan).ToolTipText = Get_Caption(waScrToolTip, tcCan) & "(F7)"
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F9)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F10)"
    

End Sub


Private Sub optDocType_Click(Index As Integer)
    Call LoadRecord
End Sub

Private Sub optDocType_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call LoadRecord
        tblDetail.SetFocus
        
    End If
End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .Update
    End With
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsLotNo As String


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
           
                
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
  
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case SLOCCODE
                
                 If .Columns(SLOCCODE).Text <> "" Then
                    
                    frmEIM0011.InTrnCd = .Columns(STRNCODE).Text
                    frmEIM0011.InLocCode = .Columns(SLOCCODE).Text
                    frmEIM0011.InUsrID = .Columns(SUSRID).Text
                    frmEIM0011.InDteTim = .Columns(SDTETIM).Text
                    frmEIM0011.InStatus = IIf(Opt_Getfocus(optDocType, 2, 0) = 0, "1", "4")
                    frmEIM0011.Show vbModal
                 
                 End If
                
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
            
        Case vbKeyReturn
            Select Case .Col
            Case SQTY
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
                Case SQTY
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






Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                
                
             
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        
        Case tcOK
        
            Call cmdSave(1)
            
        Case tcCan
        
            Call cmdSave(2)
          
          
        Case tcImport
        
            Call cmdImport
        
        Case tcCancel
        
            Call cmdCancel
            
        Case tcExit
            
            Unload Me
            
        Case tcRefresh
            
            Call LoadRecord
            
        Case tcSAll
        
           Call cmdSelect(1)
        
        Case tcDAll
        
           Call cmdSelect(0)
            
    End Select
End Sub



Private Sub tblCommon_DblClick()
    
    
    wcCombo.Text = tblCommon.Columns(0).Text
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
        wcCombo.Text = tblCommon.Columns(0).Text
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
        
        For wiCtr = SSEL To SID
            .Columns(wiCtr).AllowSizing = False
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
                Case SLOCCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                 Case SLOCNAME
                   .Columns(wiCtr).Width = 4000
                   .Columns(wiCtr).DataWidth = 50
                Case SUSRID
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 20
                Case SDTETIM
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 50
                Case SQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case STRNCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 2
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
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wsSts As String
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
        wsSts = "1"
     Else
        wsSts = "4"
    End If
    
  '  wsDteTim = "2001-04-17 00:00:00.000"
    wsSql = "SELECT EIHDUSRID, EIHDDTETIM, EIHDLOCCODE, LOCNAME, COUNT(EIHDDOCNO) QTY, CASE EIHDTRNCODE WHEN 'EM' THEN 'EM' ELSE 'EO' END TRNCODE "
    wsSql = wsSql & " FROM INEIHD, MSTLOCATION "
    wsSql = wsSql & " WHERE EIHDSTATUS = '" & wsSts & "' "
    wsSql = wsSql & " AND EIHDLOCCODE = LOCCODE "
    wsSql = wsSql & " GROUP BY EIHDUSRID, EIHDDTETIM, EIHDLOCCODE, LOCNAME, CASE EIHDTRNCODE WHEN 'EM' THEN 'EM' ELSE 'EO' END "
    wsSql = wsSql & " ORDER BY EIHDUSRID, EIHDDTETIM, EIHDLOCCODE, LOCNAME "
    
     rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SSEL, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
    
     
    With waResult
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SUSRID) = ReadRs(rsRcd, "EIHDUSRID")
        waResult(.UpperBound(1), SDTETIM) = ReadRs(rsRcd, "EIHDDTETIM")
        waResult(.UpperBound(1), SLOCCODE) = ReadRs(rsRcd, "EIHDLOCCODE")
        waResult(.UpperBound(1), SLOCNAME) = ReadRs(rsRcd, "LOCNAME")
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), STRNCODE) = ReadRs(rsRcd, "TRNCODE")
        waResult(.UpperBound(1), SID) = To_Value(ReadRs(rsRcd, "EIHDDOCID"))
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


Private Function Chk_GrdRow(ByVal LastRow As Long) As Boolean

    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    Dim OutISBN As String
    
    Chk_GrdRow = False
    
    On Error GoTo Chk_GrdRow_Err
    
    With tblDetail
        
        If To_Value(LastRow) > waResult.UpperBound(1) Then
           Chk_GrdRow = True
           Exit Function
        End If
        
        
         If Chk_EiItemNotExist(waResult(LastRow, SUSRID), waResult(LastRow, SDTETIM), OutISBN) Then
              gsMsg = OutISBN & " 沒有資料!"
              MsgBox gsMsg, vbOKOnly, gsTitle
              .Col = SSEL
              .Row = LastRow
         Exit Function
         End If
         
        
         If waResult(LastRow, SQTY) <= 0 Then
              gsMsg = "沒有數量!"
              MsgBox gsMsg, vbOKOnly, gsTitle
              .Col = SQTY
              .Row = LastRow
         Exit Function
         End If
         
         
      
     
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow " & Err.Description
    
End Function




Private Sub cmdSave(ByVal wiActFlg As Integer)

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim i As Integer
    Dim wsDocNo As String
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
     
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If
    '' Last Check when Add
    If wiActFlg = 1 Then
        gsMsg = "你是否確認匯入?"
        wiActFlg = IIf(wsTrnCd = "EM", 3, 1)
    Else
    gsMsg = "你是否刪除匯入?"
    End If
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
   i = 1
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 

    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_EIM001A"
        adcmdSave.CommandType = adCmdStoredProc
        
        'Added by Lewis at 08262002
        adcmdSave.Properties.Item("Command Time Out").Value = giTimeOut
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wiActFlg)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, SUSRID))
                Call SetSPPara(adcmdSave, 3, Change_SQLDate(waResult(wiCtr, SDTETIM)))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, SLOCCODE))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, STRNCODE))
                Call SetSPPara(adcmdSave, 6, IIf(i = wiActRow, "Y", "N"))
                Call SetSPPara(adcmdSave, 7, wsFormID)
                Call SetSPPara(adcmdSave, 8, gsUserID)
                Call SetSPPara(adcmdSave, 9, gsSystemDate)
                Call SetSPPara(adcmdSave, 10, gsLangID)
                adcmdSave.Execute
                wsDocNo = GetSPPara(adcmdSave, 11)
                If wsDocNo = "-1" Then
                    GoTo cmdSave_ItemErr
                End If
                i = i + 1
                
            End If
        Next
    End If
    

    
    cnCon.CommitTrans
    
    gsMsg = wsDocNo & " 已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Call LoadRecord
    Set adcmdSave = Nothing
    
    
    MousePointer = vbDefault
    

    
    Exit Sub
    
cmdSave_ItemErr:
    gsMsg = "書本不存在!不能完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
    
    
End Sub

Private Function InputValidation() As Boolean
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    

    wiActRow = 0
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SSEL)) = "-1" Then
                wiActRow = wiActRow + 1
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
                wsTrnCd = Trim(waResult(wlCtr, STRNCODE))
                
                
                For wlCtr1 = 0 To .UpperBound(1)
                If Trim(waResult(wlCtr1, SSEL)) = "-1" And wlCtr <> wlCtr1 Then
                If Trim(waResult(wlCtr, STRNCODE)) <> Trim(waResult(wlCtr1, STRNCODE)) Then
                    gsMsg = "不能匯入不同類別!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    Exit Function
                End If
                End If
                Next wlCtr1
            
            
                For wlCtr1 = 0 To .UpperBound(1)
                If Trim(waResult(wlCtr1, SSEL)) = "-1" And wlCtr <> wlCtr1 Then
                If Trim(waResult(wlCtr, SLOCCODE)) <> Trim(waResult(wlCtr1, SLOCCODE)) Then
                    gsMsg = "不能匯入不同地點!"
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    Exit Function
                End If
                End If
                Next wlCtr1
            
            End If
            
        Next wlCtr
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


Private Sub cmdSelect(ByVal wiSelect As Integer)
    Dim wiCtr As Long
    
    Me.MousePointer = vbHourglass
    
    
     
    With waResult
    For wiCtr = 0 To .UpperBound(1)
        waResult(wiCtr, SSEL) = IIf(wiSelect = 1, "-1", "0")
    Next wiCtr
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    Me.MousePointer = vbNormal
    
End Sub

Private Sub txtFilePath_GotFocus()
FocusMe txtFilePath
End Sub

Private Sub txtFilePath_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_txtFilePath = False Then
            Exit Sub
        End If
        
        cmdFilePath.SetFocus
        
    End If
End Sub

Private Sub txtFilePath_LostFocus()
FocusMe txtFilePath, True
End Sub

Private Function chk_txtFilePath() As Boolean

    chk_txtFilePath = False
    
     If Trim(txtFilePath) = "" Then
        gsMsg = "請輸入路徑!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath.SetFocus
        Exit Function
    End If
    
    If Not Trim(txtFilePath.Text) Like "*" & Dir(txtFilePath) & "*" Then
        gsMsg = "找不到檔案!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath.SetFocus
        Exit Function
    End If

    chk_txtFilePath = True

End Function

Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property


Public Function Chk_EiItemNotExist(ByVal InUserID As String, ByVal InDteTim As String, ByRef OutDocNo As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_EiItemNotExist = False
    OutDocNo = ""
    
    wsSql = "SELECT DISTINCT EIDTISBN FROM INEIHD, INEIDT "
    wsSql = wsSql & " Where EIHDDOCID = EIDTDOCID "
    wsSql = wsSql & " And EIHDUSRID = '" & InUserID & "'"
    wsSql = wsSql & " And EIHDDTETIM = '" & Change_SQLDate(InDteTim) & "'"
    wsSql = wsSql & " And EIHDSTATUS = '1'"
    wsSql = wsSql & " And NOT EXISTS (SELECT NULL FROM MSTITEM WHERE ITMCODE = EIDTISBN "
    wsSql = wsSql & " AND ITMSTATUS = '1')"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        rsRcd.MoveFirst
        Chk_EiItemNotExist = True
        Do While Not rsRcd.EOF
        If OutDocNo = "" Then
            OutDocNo = ReadRs(rsRcd, "EIDTISBN")
        Else
            OutDocNo = OutDocNo & ", " & ReadRs(rsRcd, "EIDTISBN")
        End If
        rsRcd.MoveNext
        Loop
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
            
    
    
End Function
