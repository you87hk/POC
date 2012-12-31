VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHHIM001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnReceive 
      Caption         =   "Receive"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox HHList 
      Height          =   240
      Left            =   10440
      TabIndex        =   10
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame fraSelect 
      Height          =   525
      Left            =   120
      TabIndex        =   9
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
         TabIndex        =   1
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
         TabIndex        =   0
         Top             =   200
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
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
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   19
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
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   9960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":1D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":208C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":24DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":2C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":2F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":33B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":3C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":3FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":440E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":472A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":4A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":4E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":51B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":54D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":57F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":5C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHHIM001.frx":5F66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   7095
      Left            =   120
      OleObjectBlob   =   "frmHHIM001.frx":6280
      TabIndex        =   6
      Top             =   1080
      Width           =   11655
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   9600
      TabIndex        =   5
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   345
      Left            =   9000
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   345
      Left            =   6600
      TabIndex        =   11
      Top             =   600
      Width           =   1290
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmHHIM001"
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
Private wsDteTim As String


Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wsTrnCd As String
Private wiActRow As Integer

Private wsHHPath As String



Private Const tcUpdate = "Update"
Private Const tcDelete = "Delete"


Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const SUSRID = 1
Private Const SDTETIM = 2
Private Const SDATE = 3
Private Const SHH = 4
Private Const SDOCNO = 5
Private Const SQTY = 6
Private Const SDUMMY = 7
Private Const SID = 8




Private Sub btnImport_Click()
    Call cmdImport
    
End Sub

Private Sub btnReceive_Click()

  MousePointer = vbHourglass
    
  Call ReceiveFromHH(wsHHPath)
  
  MsgBox "End of Receiving"
  
  Call LoadHHList
  
  
MousePointer = vbNormal

End Sub

Private Sub LoadHHList()

   
  Dim MyFile
  
   HHList.Clear
    
   MyFile = Dir(wsHHPath, vbNormal)
   Do While MyFile <> ""
    
    HHList.AddItem MyFile
    
    MyFile = Dir
   Loop
   
   btnImport.Enabled = (HHList.ListCount > 0)
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
Dim sFullFile As String
Dim sFileName As String
Dim sExt As String
Dim SDOCNO As String
Dim i As Integer
Dim bPost As Boolean
Dim relUpdFlg As String

Dim sfile As String
Dim wsDteTim As String
Dim sBKFile As String


On Error GoTo cmdImport_Err
     
  MousePointer = vbHourglass
  
  wsDteTim = Change_SQLDate(Now)
  gsMsg = ""

For i = 0 To HHList.ListCount - 1 'hidden ListBox
   ' sFullFile = HHList.List(i)
   ' sFileName = Right(sFullFile, Len(sFullFile) - InStrRev(sFullFile, "\"))
   
    sFileName = HHList.List(i)
    sFullFile = wsHHPath & HHList.List(i)
    sExt = Right(sFileName, Len(sFileName) - InStr(sFileName, "."))

    SDOCNO = Left(sFileName, InStr(sFileName, ".") - 1) & sExt
    
    sBKFile = SDOCNO & "_" & Format(Now, "YYYYMMDDHHMMSS") & "." & sExt
    
    
    If UCase(sExt) <> "BAK" And UCase(sExt) <> "FLD" Then
    
    If Chk_HHNo(SDOCNO, relUpdFlg) = True Then
        If relUpdFlg = "N" Then
        
            gsMsg = SDOCNO & "已匯入但未更新, 你是否確認要覆寫此文件?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
            bPost = False
            Else
            bPost = True
            End If
        
        Else
            gsMsg = SDOCNO & " 已匯入並更新!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            bPost = False
           ' Name sFullFile As Left(sFullFile, Len(sFullFile) - 3) & "BAK"
        End If
    Else
        bPost = True
    End If
    
    If bPost Then
        If ImportFromHH(gsUserID, wsDteTim, sFullFile) = True Then
        gsMsg = gsMsg & IIf(gsMsg = "", SDOCNO, Chr(10) & Chr(13) & SDOCNO)
        End If
    End If
    
    End If

'    Name sFullFile As wsHHPath & "backup\" & sFileName
    Name sFullFile As wsHHPath & "backup\" & sBKFile
    
Next i


 
 gsMsg = "匯入完成!"
 MsgBox gsMsg, vbOKOnly, gsTitle
 
 Call LoadHHList
 Call LoadRecord
 
 MousePointer = vbDefault
 
 Exit Sub
 
cmdImport_Err:
 MousePointer = vbDefault
 MsgBox Err.Description
 

    
    
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
    
    
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    
    wiExit = False
    wsTrnCd = ""
    HHList.Visible = False
    medPrdFr.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
    medPrdTo.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
    
    
    
    If Trim(gsHHPath) <> "" Then
    wsHHPath = gsHHPath + "receive\"
    Else
    wsHHPath = App.Path + "receive\"
    End If
     
      
    Call LoadHHList
      
    Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmHHIM001 = Nothing


    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    optDocType(0).Value = True
    wsFormID = "HHIM001"
    
     
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    

    optDocType(0).Caption = Get_Caption(waScrItm, "OPT1")
    optDocType(1).Caption = Get_Caption(waScrItm, "OPT2")
    
    btnReceive.Caption = Get_Caption(waScrItm, "RECEIVE")
    btnImport.Caption = Get_Caption(waScrItm, "IMPORT")
    
    
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    
        
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SUSRID).Caption = Get_Caption(waScrItm, "SUSRID")
        .Columns(SDTETIM).Caption = Get_Caption(waScrItm, "SDTETIM")
        .Columns(SDATE).Caption = Get_Caption(waScrItm, "SDATE")
        .Columns(SHH).Caption = Get_Caption(waScrItm, "SHH")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        
    End With
    
    
    tbrProcess.Buttons(tcUpdate).ToolTipText = Get_Caption(waScrItm, tcUpdate) & "(F6)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrItm, tcDelete) & "(F7)"
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrItm, tcSAll) & "(F9)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrItm, tcDAll) & "(F10)"
    

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
        .UPDATE
    End With
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error GoTo tblDetail_BeforeColUpdate_Err
    

       
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
                
                 If .Columns(SDOCNO).Text <> "" Then
                    
                    
                    
                    frmHHIM0011.InDocID = .Columns(SID).Text
                    frmHHIM0011.TrnCd = wsTrnCd
                    frmHHIM0011.inDteTim = .Columns(SDTETIM).Text
                    frmHHIM0011.FormID = "HHIM0011"
                    frmHHIM0011.Show vbModal
                    
                    
                    
                    
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
        
        Case tcUpdate
        
            Call cmdSave(1)
            
        Case tcDelete
        
            Call cmdSave(2)
          
          

        
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
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).ValueItems.Presentation = dbgCheckBox
                    .Columns(wiCtr).Locked = False
                Case SHH
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 10
                 Case SDATE
                   .Columns(wiCtr).Width = 1500
                   .Columns(wiCtr).DataWidth = 50
                Case SUSRID
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 20
                Case SDTETIM
                    .Columns(wiCtr).Width = 2500
                    .Columns(wiCtr).DataWidth = 50
                Case SQTY
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case SDOCNO
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Button = True
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
                Case SID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 20
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wsSts As String
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
        wsSts = "N"
     Else
        wsSts = "Y"
    End If
    
  '  wsDteTim = "2001-04-17 00:00:00.000"
    wsSQL = "SELECT HHUSRID, HHDTETIM, HHDATE, HHTID, HHNO, COUNT(HHNO) QTY, HHTYPE"
    wsSQL = wsSQL & " FROM SYSHHIM001 "
    wsSQL = wsSQL & " WHERE HHUPDFLG = '" & wsSts & "' "
    
    wsSQL = wsSQL & " AND HHDATE BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "0000/00/00"), Left(medPrdFr.Text, 4) & "/" & Right(medPrdFr.Text, 2)) & "/01" & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "9999/99/99"), Left(medPrdTo.Text, 4) & "/" & Right(medPrdTo.Text, 2)) & "/31" & "'"
    
    
    wsSQL = wsSQL & " GROUP BY HHUSRID, HHDTETIM, HHDATE, HHTID, HHNO, HHTYPE "
    
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
     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SUSRID) = ReadRs(rsRcd, "HHUSRID")
        waResult(.UpperBound(1), SDTETIM) = ReadRs(rsRcd, "HHDTETIM")
        waResult(.UpperBound(1), SHH) = ReadRs(rsRcd, "HHTID")
        waResult(.UpperBound(1), SDATE) = ReadRs(rsRcd, "HHDATE")
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "HHNO")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "HHNO")
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
    Dim wsErr As String
    Dim wsRtn As String
    Dim wsDoc As String
    
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
     
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If
    '' Last Check when Add
    If wiActFlg = 1 Then
        gsMsg = "你是否確認更新?"
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
        adcmdSave.CommandText = "USP_HHIM001A"
        adcmdSave.CommandType = adCmdStoredProc
        
        'Added by Lewis at 08262002
       ' adcmdSave.Properties.Item("Command Time Out").Value = giTimeOut
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wiActFlg)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, SDOCNO))
                Call SetSPPara(adcmdSave, 3, wsFormID)
                Call SetSPPara(adcmdSave, 4, gsUserID)
                Call SetSPPara(adcmdSave, 5, gsSystemDate)
                adcmdSave.Execute
                wsErr = GetSPPara(adcmdSave, 6)
                wsRtn = GetSPPara(adcmdSave, 7)
                If wsErr = "-1" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Such Document No: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-2" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No REMAIN QTY in: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-3" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Such Item No: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-4" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Insufficient stock: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-5" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Greater then remaining Qty: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-6" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Such Staff Code: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-7" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Such Hand-Held ID: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-8" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Need to approve picking: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-9" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Different qty with Job: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-10" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Already exists in Stock Take Doc.: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                ElseIf wsErr = "-11" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Insufficent stock from C to B: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-12" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Insufficent stock from B to A: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-13" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Must specific Bin No: " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-14" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Already upated " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-15" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Transfer Out Item, cannot transfer In~! " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-16" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " Different Qty from Transfer Out! " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                ElseIf wsErr = "-17" Then
                    gsMsg = waResult(wiCtr, SDOCNO) & " No Active Counting Period at " & wsRtn
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    GoTo cmdSave_Err
                End If
                
                
                wsDoc = IIf(wsDoc = "", wsRtn, wsDoc & "," & wsRtn)
                i = i + 1
                
            End If
        Next
    End If
    

    
    cnCon.CommitTrans
    
    gsMsg = wsDoc & "已完成!"
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
    gsMsg = "更新不能完成!"
    MsgBox gsMsg & " " & Err.Description
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
                wsTrnCd = Trim(waResult(wlCtr, SDOCNO))
                
                

            
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

Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property

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
