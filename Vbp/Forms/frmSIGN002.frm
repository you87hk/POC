VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIGN002 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmSIGN002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmSIGN002.frx":0442
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame fraSelect 
      Height          =   810
      Left            =   9000
      TabIndex        =   15
      Top             =   360
      Width           =   2775
      Begin VB.OptionButton optDocType 
         Caption         =   "SN"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDocType 
         Caption         =   "SO"
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   8895
      Begin VB.OptionButton optInOut 
         Caption         =   "IN"
         Height          =   495
         Index           =   5
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optInOut 
         Caption         =   "IN"
         Height          =   495
         Index           =   4
         Left            =   6096
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optInOut 
         Caption         =   "IN"
         Height          =   495
         Index           =   3
         Left            =   4632
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optInOut 
         Caption         =   "IN"
         Height          =   495
         Index           =   2
         Left            =   3168
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optInOut 
         Caption         =   "OUT"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optInOut 
         Caption         =   "IN"
         Height          =   495
         Index           =   1
         Left            =   1704
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "frmSIGN002.frx":2B45
      TabIndex        =   10
      Top             =   1680
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
            Picture         =   "frmSIGN002.frx":ACB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":B592
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":BE6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":C2BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":C710
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":CA2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":CE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":D2CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":D5E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":D902
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":DD54
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":E630
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":E958
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":EDAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":F0C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":F3E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":F838
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":FB54
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":FE70
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":102C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIGN002.frx":105E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   13
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
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sign"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Can"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   2520
      TabIndex        =   16
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIGN002"
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


Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer
Private wsTrnCd As String

Private wiSort As Integer
Private wsSortBy As String

Private Const tcSign = "Sign"
Private Const tcCan = "Can"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"

Private Const SSEL = 0
Private Const SDOCNO = 1
Private Const SCCODE = 2
Private Const SDOCDATE = 3
Private Const SJOBNO = 4
Private Const SCTLPRD = 5
Private Const SSRCCODE = 6
Private Const SPFX = 7
Private Const SJOURNO = 8
Private Const SDUMMY = 9
Private Const SID = 10



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF10
         If tbrProcess.Buttons(tcSign).Enabled = False Then Exit Sub
           Call cmdSave(1)
            
        Case vbKeyF3
         If tbrProcess.Buttons(tcCan).Enabled = False Then Exit Sub
  
           Call cmdSave(2)
           
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
            
        Case vbKeyF5
           Call cmdSelect(1)
           
        Case vbKeyF6
           Call cmdSelect(0)
        
        Case vbKeyF7
            Call LoadRecord
        
      
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
    
    optDocType(0).Value = True
    optInOut(0).Value = True
    
    wiSort = 0
    wsSortBy = "ASC"
    
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    
    
    medPrdFr.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
    medPrdTo.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))

    
    
    Call LoadRecord
     
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmSIGN002 = Nothing

    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "SIGN002"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
       
    optInOut(0).Caption = Get_Caption(waScrItm, "OPT0")
    optInOut(1).Caption = Get_Caption(waScrItm, "OPT1")
    optInOut(2).Caption = Get_Caption(waScrItm, "OPT2")
    optInOut(3).Caption = Get_Caption(waScrItm, "OPT3")
    optInOut(4).Caption = Get_Caption(waScrItm, "OPT4")
    optInOut(5).Caption = Get_Caption(waScrItm, "OPT5")
    
    optDocType(0).Caption = Get_Caption(waScrItm, "STS1")
    optDocType(1).Caption = Get_Caption(waScrItm, "STS2")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCCODE).Caption = Get_Caption(waScrItm, "SCCODE")
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SJOBNO).Caption = Get_Caption(waScrItm, "SJOBNO")
        .Columns(SCTLPRD).Caption = Get_Caption(waScrItm, "SCTLPRD")
        .Columns(SSRCCODE).Caption = Get_Caption(waScrItm, "SSRCCODE")
        .Columns(SPFX).Caption = Get_Caption(waScrItm, "SPFX")
        .Columns(SJOURNO).Caption = Get_Caption(waScrItm, "SJOURNO")
        
       
    End With
    
    
    tbrProcess.Buttons(tcSign).ToolTipText = Get_Caption(waScrToolTip, tcSign) & "(F10)"
    tbrProcess.Buttons(tcCan).ToolTipText = Get_Caption(waScrToolTip, tcCan) & "(F3)"
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F5)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F6)"
    
    

End Sub



Private Sub optDocType_Click(Index As Integer)
  Call LoadRecord
End Sub

Private Sub optInOut_Click(Index As Integer)
   Call LoadRecord
End Sub

Private Sub optInOut_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
       Call LoadRecord
        
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
            Case SDOCNO
                
           '      If .Columns(SDOCNO).Text <> "" Then
           '         Call getTrnCd
           '
           '         frmSIGN0011.InDocID = .Columns(SID).Text
           '         frmSIGN0011.InCusNo = .Columns(STRNCODE).Text
           '         frmSIGN0011.InTrnCd = wsTrnCd
           '         frmSIGN0011.Show vbModal
           '
           '      End If
                
           End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDeiail ButtonClick!"
 
End Sub

Private Sub tblDetail_HeadClick(ByVal ColIndex As Integer)

    
    On Error GoTo tblDetail_HeadClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case SDOCNO
                wiSort = 0
                cmdRefresh
            Case SDOCDATE
                wiSort = 1
                cmdRefresh
            Case SCCODE
                wiSort = 2
                cmdRefresh
           End Select
    End With

    
    Exit Sub
    
tblDetail_HeadClick_Err:
     MsgBox "Check tblDeiail HeadClick!"

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
            Case SJOURNO
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
                Case SJOURNO
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
            Case SCCODE
                If Opt_Getfocus(optInOut, 6, 0) < 3 Then
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(.Columns(SCCODE).Text) & "'", "CUSNAME")
                Else
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = Get_TableInfo("MSTVENDOR", "VDRCODE = '" & Set_Quote(.Columns(SCCODE).Text) & "'", "VDRNAME")
                
                End If
                 
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)

  If tbrProcess.Buttons(Button.Key).Enabled = False Then Exit Sub
 
    Select Case Button.Key
        Case tcSign
            Call cmdSave(1)
            
        Case tcCan
            Call cmdSave(2)
            
        Case tcCancel
        
           Call cmdCancel
           
        Case tcSAll
           Call cmdSelect(1)
           
        Case tcDAll
           Call cmdSelect(0)
            
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call LoadRecord
            
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
    
    
 On Error GoTo tblCommon_LostFocus_Err
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If
    
Exit Sub
tblCommon_LostFocus_Err:

Set wcCombo = Nothing

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
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).Locked = False
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                Case SCCODE
                   .Columns(wiCtr).Width = 1000
                   .Columns(wiCtr).DataWidth = 10
                Case SDOCDATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SJOBNO
                   .Columns(wiCtr).Width = 2500
                   .Columns(wiCtr).DataWidth = 20
                Case SCTLPRD
                    .Columns(wiCtr).Width = 800
                     .Columns(wiCtr).DataWidth = 6
                Case SSRCCODE
                    .Columns(wiCtr).Width = 1000
                     .Columns(wiCtr).DataWidth = 20
                 Case SPFX
                    .Columns(wiCtr).Width = 1000
                     .Columns(wiCtr).DataWidth = 10
                 Case SJOURNO
                    .Columns(wiCtr).Width = 1500
                     .Columns(wiCtr).DataWidth = 10
                     .Columns(wiCtr).Locked = False
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
    Dim wsUpdFlg As String
    Dim wsPfx As String
    Dim wsJourNo As String
    
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    Call Set_tbrProcess
    Call getTrnCd
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
     wsUpdFlg = "N"
    Else
     wsUpdFlg = "Y"
    End If
    
    
    
    Select Case wsTrnCd
    Case "ARA"
    
    wsSQL = "SELECT INHDDOCID DOCID, INHDDOCNO DOCNO, INHDDOCDATE DOCDATE, CUSCODE CCODE, INHDJOBNO JOBNO, INHDCTLPRD CTLPRD, 'ARA' SRCCODE, INHDVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  ARINHD, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE INHDCUSID = CUSID "
    wsSQL = wsSQL & " AND INHDUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND INHDPGMNO <> 'OPN' "
    wsSQL = wsSQL & " AND INHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY INHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY INHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY INHDDOCNO, INHDDOCDATE, CUSCODE " & wsSortBy
    End If
    
    
    Case "ARB"
    
    wsSQL = "SELECT ARCQCHQID DOCID, ARCQCHQNO DOCNO, ARCQCHQDATE DOCDATE, CUSCODE CCODE, '' JOBNO, ARCQCTLPRD CTLPRD, 'ARB' SRCCODE, ARCQVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  ARCHEQUE, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE ARCQCUSID = CUSID "
    wsSQL = wsSQL & " AND ARCQUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND ARCQPGMNO = 'AR002' "
    wsSQL = wsSQL & " AND ARCQCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY ARCQCHQNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY ARCQCHQDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY ARCQCHQNO, ARCQCHQDATE, CUSCODE " & wsSortBy
    End If
    
  Case "ARC"
    
    wsSQL = "SELECT ARSHDOCID DOCID, ARSHDOCNO DOCNO, ARSHDOCDATE DOCDATE, CUSCODE CCODE, '' JOBNO, ARSHCTLPRD CTLPRD, 'ARC' SRCCODE, ARSHVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  ARSTHD, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE ARSHCUSID = CUSID "
    wsSQL = wsSQL & " AND ARSHUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND ARSHCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY ARSHDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY ARSHDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY ARSHDOCNO, ARSHDOCDATE, CUSCODE " & wsSortBy
    End If
    
   Case "APA"
    
    wsSQL = "SELECT IPHDDOCID DOCID, IPHDDOCNO DOCNO, IPHDDOCDATE DOCDATE, VDRCODE CCODE, IPHDJOBNO JOBNO, IPHDCTLPRD CTLPRD, 'APA' SRCCODE, IPHDVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  APIPHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE IPHDVDRID = VDRID "
    wsSQL = wsSQL & " AND IPHDUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND IPHDPGMNO <> 'OPN' "
    wsSQL = wsSQL & " AND IPHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY IPHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY IPHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY VDRCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY IPHDDOCNO, IPHDDOCDATE, VDRCODE " & wsSortBy
    End If
    
    
    Case "APB"
    
    wsSQL = "SELECT APCQCHQID DOCID, APCQCHQNO DOCNO, APCQCHQDATE DOCDATE, VDRCODE CCODE, '' JOBNO, APCQCTLPRD CTLPRD, 'APB' SRCCODE, APCQVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  APCHEQUE, MSTVENDOR "
    wsSQL = wsSQL & " WHERE APCQVDRID = VDRID "
    wsSQL = wsSQL & " AND APCQUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND APCQPGMNO = 'AP002' "
    wsSQL = wsSQL & " AND APCQCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY APCQCHQNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY APCQCHQDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY VDRCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY APCQCHQNO, APCQCHQDATE, VDRCODE " & wsSortBy
    End If
    
  Case "APC"
    
    wsSQL = "SELECT APSHDOCID DOCID, APSHDOCNO DOCNO, APSHDOCDATE DOCDATE, VDRCODE CCODE, '' JOBNO, APSHCTLPRD CTLPRD, 'APC' SRCCODE, APSHVOUCNO JOURNO "
    wsSQL = wsSQL & " FROM  APSTHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE APSHVDRID = VDRID "
    wsSQL = wsSQL & " AND APSHUPDFLG = '" & wsUpdFlg & "' "
    wsSQL = wsSQL & " AND APSHCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY APSHDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY APSHDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY VDRCODE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY APSHDOCNO, APSHDOCDATE, VDRCODE " & wsSortBy
    End If
    
    
    
    End Select
    
    
     rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

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

    wsPfx = Get_TableInfo("SYSVOUNO", "VOUTYPE = '" & Left(ReadRs(rsRcd, "SRCCODE"), 2) & "'", "VOUPREFIX")
    If IsNull(ReadRs(rsRcd, "JOURNO")) Or ReadRs(rsRcd, "JOURNO") = "" Then
    wsJourNo = ""
    Else
    wsJourNo = Right(ReadRs(rsRcd, "JOURNO"), Len(ReadRs(rsRcd, "JOURNO")) - 4)
    End If

     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), SCCODE) = ReadRs(rsRcd, "CCODE")
        waResult(.UpperBound(1), SJOBNO) = ReadRs(rsRcd, "JOBNO")
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "DOCDATE")
        waResult(.UpperBound(1), SCTLPRD) = ReadRs(rsRcd, "CTLPRD")
        waResult(.UpperBound(1), SSRCCODE) = ReadRs(rsRcd, "SRCCODE")
        waResult(.UpperBound(1), SPFX) = wsPfx
'        waResult(.UpperBound(1), SJOURNO) = IIf(ReadRs(rsRcd, "JOURNO") <> "", Right(ReadRs(rsRcd, "JOURNO"), Len(ReadRs(rsRcd, "JOURNO")) - 4), "")
         waResult(.UpperBound(1), SJOURNO) = wsJourNo
         
       
         waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DOCID")
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
    
    Chk_GrdRow = False
    
    On Error GoTo Chk_GrdRow_Err
    
    With tblDetail
        
        If To_Value(LastRow) > waResult.UpperBound(1) Then
           Chk_GrdRow = True
           Exit Function
        End If
        
        
        If wiActFlg = 1 Then
        If Chk_grdJourNo(waResult(LastRow, SPFX), waResult(LastRow, SJOURNO)) = False Then
                .Col = SJOURNO
                Exit Function
        End If
        End If
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function




Private Sub cmdSave(ByVal inActFlg As Integer)

    Dim wsGenDte As String
    Dim wsDteTim As String
    
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    wsDteTim = Change_SQLDate(Now)
    
    wiActFlg = inActFlg
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If
    '' Last Check when Add
   
    Select Case wiActFlg
    Case 1
    gsMsg = "你是否確認此文件?"
    Case 2
    gsMsg = "你是否取消此批核?"
    End Select
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
    Call getTrnCd
       
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
    
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_ACC000A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wiActFlg)
                Call SetSPPara(adcmdSave, 2, wsTrnCd)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, SDOCNO))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, SCCODE))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, SDOCDATE))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, SJOURNO))
                Call SetSPPara(adcmdSave, 8, gsUserID)
                Call SetSPPara(adcmdSave, 9, wsGenDte)
                Call SetSPPara(adcmdSave, 10, wsDteTim)
                adcmdSave.Execute
            End If
        Next
    End If
    

    
    cnCon.CommitTrans
    
    gsMsg = "已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Call LoadRecord
    Set adcmdSave = Nothing
    
    
    MousePointer = vbDefault
    
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
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SSEL)) = "-1" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
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


Private Sub getTrnCd()

Select Case Opt_Getfocus(optInOut, 6, 0)

Case 0
    wsTrnCd = "ARA"
Case 1
    wsTrnCd = "ARB"
Case 2
    wsTrnCd = "ARC"
Case 3
    wsTrnCd = "APA"
Case 4
    wsTrnCd = "APB"
Case 5
    wsTrnCd = "APC"
End Select


End Sub

Private Sub Set_tbrProcess()

With tbrProcess
    
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
    .Buttons(tcCan).Enabled = False
    .Buttons(tcSign).Enabled = True
    Else
    .Buttons(tcCan).Enabled = True
    .Buttons(tcSign).Enabled = False
    End If
    
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    
End With

End Sub


Private Sub cmdRefresh()


    If wsSortBy = "ASC" Then
    wsSortBy = "DESC"
    Else
    wsSortBy = "ASC"
    End If
    LoadRecord
    
End Sub


Private Function Chk_grdJourNo(ByVal InPfx As String, ByVal inAccNo As String) As Boolean
Dim wsStatus As String
    
     Chk_grdJourNo = False
    
    
    If Trim(inAccNo) = "" Then
        Chk_grdJourNo = True
        Exit Function
    End If
    
    
   If Chk_VoucNo(InPfx, inAccNo) = True Then
        
       gsMsg = "憑單編號已用,不能再用!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
    Chk_grdJourNo = True
    
    
End Function
Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub


Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        
        If LoadRecord = True Then
            tblDetail.SetFocus
        End If
       
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub
Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function
Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        gsMsg = "To must > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function


