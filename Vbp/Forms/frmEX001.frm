VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEX001 
   Caption         =   "Out Sales Data For Generation PO"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9270
   Icon            =   "frmEX001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   9270
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   9000
      OleObjectBlob   =   "frmEX001.frx":030A
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2160
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5190
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5190
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2160
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      ScaleHeight     =   360
      ScaleWidth      =   8505
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   8565
   End
   Begin VB.Frame fmeFilePath 
      Caption         =   "Output Path"
      Height          =   780
      Index           =   1
      Left            =   312
      TabIndex        =   7
      Top             =   2856
      Width           =   8604
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Left            =   1848
         TabIndex        =   3
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   288
         Width           =   6312
      End
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Left            =   8160
         TabIndex        =   4
         Top             =   288
         Width           =   288
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Export Sheet Path:"
         Height          =   288
         Left            =   168
         TabIndex        =   8
         Top             =   336
         Width           =   1476
      End
   End
   Begin MSComDlg.CommonDialog cdFilePath 
      Left            =   0
      Top             =   840
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   -120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEX001.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fmeFilePath 
      Height          =   2310
      Index           =   0
      Left            =   312
      TabIndex        =   6
      Top             =   405
      Width           =   8604
      Begin MSMask.MaskEdBox medETA 
         Height          =   330
         Left            =   1845
         TabIndex        =   2
         Top             =   1845
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOrdPrdTo 
         Height          =   288
         Left            =   4908
         TabIndex        =   1
         Top             =   288
         Width           =   1392
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOrdPrdFr 
         Height          =   288
         Left            =   1848
         TabIndex        =   0
         Top             =   288
         Width           =   1392
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medWant 
         Height          =   330
         Left            =   1845
         TabIndex        =   13
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label lblCusNoTo 
         Caption         =   "To"
         Height          =   345
         Left            =   4440
         TabIndex        =   24
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblDocNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4440
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCusNoFr 
         Caption         =   "Customer Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1890
      End
      Begin VB.Label lblWant 
         Caption         =   "WANTDATE"
         Height          =   225
         Left            =   165
         TabIndex        =   14
         Top             =   1485
         Width           =   1665
      End
      Begin VB.Label lblETA 
         Caption         =   "ETADATE"
         Height          =   225
         Left            =   165
         TabIndex        =   11
         Top             =   1890
         Width           =   1665
      End
      Begin VB.Label lblOrdPrdTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4440
         TabIndex        =   10
         Top             =   330
         Width           =   375
      End
      Begin VB.Label lblOrdPrdFr 
         Caption         =   "Order Period From:"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Progress:"
      Height          =   252
      Left            =   456
      TabIndex        =   5
      Top             =   3768
      Width           =   8376
   End
End
Attribute VB_Name = "frmEX001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim wsFormID As String
Dim wcCombo As Control
Private wsFormCaption As String
Private wsExcelPath As String

Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB

Dim wsGenDte As String

'Frame
Private Const inPath = 0
Private Const OutPath = 1
'P/O Data Upload
Private Const DocNo = 1
Private Const CusCode = 2
Private Const CusName = 3
Private Const DocDate = 4
Private Const RevNo = 5
Private Const Curr = 6
Private Const Excr = 7
Private Const DueDate = 8
Private Const EtaDate = 9
Private Const Ln = 10
Private Const VdrCode = 11
Private Const ItmCode = 12
Private Const WhsCode = 13
Private Const ItmNam = 14
Private Const ItPub = 15
Private Const ItWanted = 16
Private Const Qty = 17
Private Const Uom = 18
Private Const NeedQty = 19
Private Const PoUom = 20
Private Const Price = 21
Private Const DisPer = 22
Private Const Amt = 23
Private Const Amtl = 24
Private Const Dis = 25
Private Const Disl = 26
Private Const Net = 27
Private Const Netl = 28


Private Const tcGo = "Go"
Private Const tcExit = "Exit"
Private Const tcCancel = "Cancel"


Private Sub cmdCancel()
    MousePointer = vbDefault
    Ini_Scr
    medOrdPrdFr.SetFocus
End Sub

Private Sub cmdOK()

    MousePointer = vbHourglass
    
    On Error GoTo cmdOK_Err
    
    If InputValidation = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    If OutputPoData = True Then
        gsMsg = "Execl 文件已製作完成 !"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    End If
    
    Call Ini_Scr
    MousePointer = vbDefault
    
    Exit Sub
    
cmdOK_Err:

    gsMsg = "cmdOK Err!"
    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    MousePointer = vbDefault

End Sub


Private Sub cmdFilePath_Click()
    
    On Error Resume Next
      
    If Trim(txtFilePath.Text) <> "" Then
        'If Dir(txtFilePath.Text) <> "" Then
            'Dialog.dirDirectory.Path = txtFilePath.Text
            cdFilePath.InitDir = wsExcelPath
        'End If
    End If
    'Dialog.Show vbModal
   ' If Trim(Dialog.Tag) <> "" Then
   '     txtFilePath.Text = Dialog.Tag
   '     If chk_txtFilePath = False Then
   '         txtFilePath.SetFocus
   '     End If
   ' End If
    With cdFilePath
    .DialogTitle = "Open A Execl File"
    .Filter = "Excel File (*.XLS)|*.XLS"
    .CancelError = True
    .FileName = vbNullString
    .ShowOpen
    If Err.Number <> cdlCancel Then
        txtFilePath.Text = .FileName
    End If
    End With
    
    On Error GoTo 0
   
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      
        Case vbKeyF9
            Call cmdOK
            
        Case vbKeyF11
            Call cmdCancel

        Case vbKeyF12
        
            Unload Me
    End Select
End Sub

Private Sub Form_Load()

    MousePointer = vbHourglass
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Scr

    MousePointer = vbDefault

End Sub
Private Sub Ini_Form()

    Me.KeyPreview = True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "EX001"
End Sub

Private Sub Ini_Scr()

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = " SELECT MAX(SoHdCtlPrd) AS CTLPER FROM soaSoHd "
    wsSql = wsSql & " WHERE SoHdStatus <> '2'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    Call SetPeriodMask(medOrdPrdFr)
    Call SetPeriodMask(medOrdPrdTo)
    Call SetDateMask(medETA)
    Call SetDateMask(medWant)
    
    medETA = Dsp_MedDate(gsSystemDate)
  
    If ReadRs(rsRcd, "CTLPER") <> "" Then
        medOrdPrdFr.Text = Left(ReadRs(rsRcd, "CTLPER"), 4) & "/" & _
                           Right(ReadRs(rsRcd, "CTLPER"), 2)
        medOrdPrdTo.Text = Left(ReadRs(rsRcd, "CTLPER"), 4) & "/" & _
                           Right(ReadRs(rsRcd, "CTLPER"), 2)
    End If
    
    
   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   
   
   If InStr(gsExcPath, ":\") Or InStr(gsExcPath, "\\") Then
        wsExcelPath = gsExcPath
    Else
        wsExcelPath = App.Path & "\" & gsExcPath
    End If
    
    txtFilePath.Text = wsExcelPath & "SC_" & _
                    Format(Now, "YYYYMMDDHHMM") & ".XLS"
   
    UpdStatusBar picStatus, 0
    lblStatus.Caption = ""
    wsGenDte = ""
    
End Sub



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 5085
        Me.Width = 9390
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmEX001 = Nothing

End Sub

Private Sub Ini_Caption()

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")

    fmeFilePath(OutPath).Caption = Get_Caption(waScrItm, "OUTPATH")
    
    lblOrdPrdFr.Caption = Get_Caption(waScrItm, "ORDPRDFR")
    lblOrdPrdTo.Caption = Get_Caption(waScrItm, "ORDPRDTO")
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

    
    lblETA.Caption = Get_Caption(waScrItm, "ETA")
    lblWant.Caption = Get_Caption(waScrItm, "WANT")
   
    lblFilePath.Caption = Get_Caption(waScrItm, "OUTPUTPATH")
    lblStatus.Caption = Get_Caption(waScrItm, "STATUS")
End Sub


Private Sub medETA_GotFocus()
    FocusMe medETA
End Sub

Private Sub medETA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medETA = False Then
            Exit Sub
        End If
            
        txtFilePath.SetFocus
    End If
End Sub

Private Sub medETA_LostFocus()
    FocusMe medETA, True
End Sub

Private Sub medWant_GotFocus()
    FocusMe medWant
End Sub

Private Sub medWant_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medWant = False Then
            Exit Sub
        End If
            
        medETA.SetFocus
    End If
End Sub

Private Sub medWant_LostFocus()
    FocusMe medWant, True
End Sub

Private Sub medOrdPrdFr_GotFocus()
    FocusMe medOrdPrdFr
End Sub

Private Sub medOrdPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medOrdPrdFr = False Then
            Exit Sub
        End If
    
        If Trim(medOrdPrdFr.Text) <> "" And _
            Trim(medOrdPrdTo.Text) = "" Then
            medOrdPrdTo.Text = medOrdPrdFr.Text
        End If
        medOrdPrdTo.SetFocus
    End If
End Sub


Private Sub medOrdPrdFr_LostFocus()
    FocusMe medOrdPrdFr, True
End Sub

Private Sub medOrdPrdTo_GotFocus()
    FocusMe medOrdPrdTo
End Sub

Private Sub medOrdPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medOrdPrdTo = False Then
            Exit Sub
        End If
        cboDocNoFr.SetFocus
    End If
End Sub


Private Sub medOrdPrdTo_LostFocus()
    FocusMe medOrdPrdTo, True
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case tcGo
            Call cmdOK
            
        Case tcCancel
        
            Call cmdCancel
            
        Case tcExit
        
            Unload Me
           
    End Select

End Sub

Private Sub txtFilePath_DblClick()
        Call LoadExcel(txtFilePath.Text)
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
        medOrdPrdFr.SetFocus

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
 '   If Dir(txtFilePath, vbDirectory) = "" Then
 '       gsMsg = "Folder Not Found!"
 '       MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
 '       txtFilePath.SetFocus
 '       Exit Function
 '   End If
    If Trim(txtFilePath.Text) Like "*" & Dir(txtFilePath) & "*" And _
        Dir(txtFilePath) <> "" Then
        gsMsg = "檔案已存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath.SetFocus
        Exit Function
    End If
    
    chk_txtFilePath = True

End Function

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_medOrdPrdFr = False Then
        medOrdPrdFr.SetFocus
        Exit Function
    End If
    
    If chk_medOrdPrdTo = False Then
        medOrdPrdTo.SetFocus
        Exit Function
    End If
            
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboCusNoTo = False Then
        cboCusNoTo.SetFocus
        Exit Function
    End If
    
    If chk_medWant = False Then
        medWant.SetFocus
        Exit Function
    End If
    
    
    If chk_medETA = False Then
        medETA.SetFocus
        Exit Function
    End If
    
    If chk_txtFilePath = False Then
        txtFilePath.SetFocus
        Exit Function
    End If
    
    InputValidation = True

End Function



Private Function chk_medOrdPrdFr() As Boolean

    chk_medOrdPrdFr = False
    
    If Trim(medOrdPrdFr.Text) = "/" Then
        chk_medOrdPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medOrdPrdFr) = False Then
        gsMsg = "不正確週期!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medOrdPrdFr.SetFocus
        Exit Function
    End If
    
    chk_medOrdPrdFr = True
    
End Function

Private Function chk_medOrdPrdTo() As Boolean

    chk_medOrdPrdTo = False
    
    If Trim(medOrdPrdTo.Text) = "/" Then
        chk_medOrdPrdTo = True
        Exit Function
    End If
    
    
    If Chk_Period(medOrdPrdTo) = False Then
        gsMsg = "不正確週期!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medOrdPrdTo.SetFocus
        Exit Function
    End If
    
    If Right(medOrdPrdFr.Text, 4) & Left(medOrdPrdFr, 2) > _
       Right(medOrdPrdTo.Text, 4) & Left(medOrdPrdTo, 2) Then
        gsMsg = "由之值不能大於至之值!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Exit Function
    End If
       
    
    chk_medOrdPrdTo = True
    
End Function

Private Function chk_medETA() As Boolean

    chk_medETA = False
    
    If Trim(medETA.Text) = "/  /" Then
        chk_medETA = True
        Exit Function
    End If
    
    
    If Chk_Date(medETA) = False Then
        gsMsg = "不正確日期!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medETA.SetFocus
        Exit Function
    End If
    
    chk_medETA = True
    
End Function

Private Function chk_medWant() As Boolean

    chk_medWant = False
    
    
    If Trim(medWant.Text) = "/  /" Then
        chk_medWant = True
        Exit Function
    End If
    
    If Chk_Date(medWant) = False Then
        gsMsg = "不正確日期!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medWant.SetFocus
        Exit Function
    End If
    
    chk_medWant = True
    
End Function


Private Sub cboCusNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_GotFocus()
        FocusMe cboCusNoFr
    Set wcCombo = cboCusNoFr
End Sub

Private Sub cboCusNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboCusNoFr.Text) <> "" And _
            Trim(cboCusNoTo.Text) = "" Then
            cboCusNoTo.Text = cboCusNoFr.Text
        End If
        cboCusNoTo.SetFocus
    End If
End Sub


Private Sub cboCusNoFr_LostFocus()
    FocusMe cboCusNoFr, True
End Sub

Private Sub cboCusNoTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoTo_GotFocus()
    FocusMe cboCusNoTo
    Set wcCombo = cboCusNoTo
End Sub

Private Sub cboCusNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo = False Then
            Exit Sub
        End If
        
        medWant.SetFocus
    End If
End Sub



Private Sub cboCusNoTo_LostFocus()
FocusMe cboCusNoTo, True
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
End Sub
Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        gsMsg = "至之值不能大於由之值"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusNoFr.SetFocus
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function
Private Function chk_cboDocNoTo() As Boolean
    chk_cboDocNoTo = False
    
    If UCase(cboDocNoFr.Text) > UCase(cboDocNoTo.Text) Then
        gsMsg = "至之值不能大於由之值"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNoFr.SetFocus
        Exit Function
    End If
    
    chk_cboDocNoTo = True
End Function
Private Sub cboDocNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSql = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSql = wsSql & " FROM soaSOHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSql = wsSql & " AND SOHDCUSID  = CUSID "
    wsSql = wsSql & " AND SOHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY SOHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboDocNoFr.Text) <> "" And _
            Trim(cboDocNoTo.Text) = "" Then
            cboDocNoTo.Text = cboDocNoFr.Text
        End If
        cboDocNoTo.SetFocus
    End If
End Sub

Private Sub cboDocNoFr_LostFocus()
    FocusMe cboDocNoFr, True
End Sub

Private Sub cboDocNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    wsSql = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSql = wsSql & " FROM soaSOHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSql = wsSql & " AND SOHDCUSID  = CUSID "
    wsSql = wsSql & " AND SOHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY SOHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Exit Sub
        End If
        
        cboCusNoFr.SetFocus
    End If
End Sub


Private Function OutputPoData() As Boolean
    
    ' Declare Database variable and Microsoft Excel Form variable.
    Dim xlApp As Excel.Application
    Dim rsTmp As New ADODB.Recordset

    
    Dim xlData As Variant
    Dim tmpData As Variant
    Dim inpParent As Variant
    Dim wsSql As String
    Dim StartRow As Long
    Dim ArrayRow As Long
    Dim i As Long
    Dim J As Long
    Dim X As Integer
    
    Dim RecordNo As Long
    Dim wiCtr As Integer
    Dim Counter As Integer
    Dim CurRow As Long
    Dim CurCol As Integer
    Dim wiStatus As Long
    Dim NoOfFty As Integer
    Dim FileName As String
    Dim dTmpValue As Double
    
    On Error GoTo OutputPoData_Err

    OutputPoData = False
    wsGenDte = Change_SQLDate(Now)
    
    UpdStatusBar picStatus, 5
    
    
    wsSql = "EXEC usp_RPTEX001 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & wsGenDte & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medOrdPrdFr.Text) = "/", String(6, "000000"), Left(medOrdPrdFr.Text, 4) & Right(medOrdPrdFr.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medOrdPrdTo.Text) = "/", String(6, "999999"), Left(medOrdPrdTo.Text, 4) & Right(medOrdPrdTo.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medWant.Text) = "/  /", "9999/99/99", medWant.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medETA.Text) = "/  /", "9999/99/99", medETA.Text) & "', "
    wsSql = wsSql & gsLangID
    cnCon.Execute wsSql
    UpdStatusBar picStatus, 10
    
    
    wiStatus = 0
    Counter = Netl
    
    'PRINT HEADING
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    UpdStatusBar picStatus, 20
    

        CurRow = 3
        StartRow = 3
        ArrayRow = 0
        xlApp.Workbooks.Open wsExcelPath & "EX001.XLS", ReadOnly:=True
        
  '      FileName = ReadRs(rsRcd, "RPTPERIOD")
  '      lblStatus.Caption = Get_Caption(waScrItm, "STATUS2") & FileName
        wsSql = " SELECT RPTDOCNO, RPTCUSCODE, RPTCUSNAME, RPTDOCDATE, RPTREVNO, "
        wsSql = wsSql & " RPTCURR, RPTEXCR, RPTDUEDATE, RPTETADATE, RPTLN, RPTVDRCODE, "
        wsSql = wsSql & " RPTITMCODE, RPTWHSCODE, RPTITMNAM, RPTITPUB, RPTWANTED, RPTQTY, "
        wsSql = wsSql & " RPTUOM, RPTNEEDQTY, RPTPOUOM, RPTPRICE, RPTDISPER, RPTAMT, RPTAMTL, RPTDIS, RPTDISL, "
        wsSql = wsSql & " RPTNET, RPTNETL "
        wsSql = wsSql & " FROM RPTEX001 "
        wsSql = wsSql & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
        wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
        wsSql = wsSql & " ORDER BY RPTDOCNO, RPTLN, RPTITMCODE "
        rsTmp.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
        RecordNo = rsTmp.RecordCount
    
        If RecordNo = 0 Then
            gsMsg = "找到不資料!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            GoTo NoDataInRpt_Err
        End If
        
        With xlApp.Sheets("Sheet1")
          '  .Cells(1, 1).Value = "'P/O DATA UPLOAD FOR PERIOD" & FileName
            ' insert details
            ReDim xlData(Counter, 0)
            
            Do While Not rsTmp.EOF
        
                wiStatus = wiStatus + Fix((1 / RecordNo) * (100))
                
                ArrayRow = ArrayRow + 1
                ReDim Preserve xlData(Counter, ArrayRow - 1)
        
                For CurCol = DocNo To Netl
                    Select Case CurCol
                    
                    Case Excr, Qty, NeedQty, Price, Amt, Amtl, Dis, Disl, Net, Netl
                        xlData(CurCol - 1, ArrayRow - 1) = To_Value(ReadRs(rsTmp, CurCol - 1))
                        
                   ' Case Pkg
                    
                    ' Tom 2000/07/05
                   '     dTmpValue = To_Value(ReadRs(rsTmp, CurCol - 1))
                   '     If dTmpValue = 0 Then
                   '     xlData(CurCol - 1, ArrayRow - 1) = ""
                   '     Else
                   '     xlData(CurCol - 1, ArrayRow - 1) = dTmpValue
                   '     End If
                        
                    Case Else
                        inpParent = ReadRs(rsTmp, CurCol - 1)
                        xlData(CurCol - 1, ArrayRow - 1) = CStr("'" & "" & inpParent)
                    End Select
                Next
            
                If ArrayRow >= 100 Or CurRow = (RecordNo + 2) Then
                    ReDim tmpData(UBound(xlData, 2), UBound(xlData, 1))
                    For i = 0 To UBound(xlData, 1)
                        For J = 0 To UBound(xlData, 2)
                            tmpData(J, i) = xlData(i, J)
                        Next
                    Next
                    .Range(.Cells(StartRow, 1).Address, .Cells(CurRow, Counter).Address).Value = tmpData
                    ArrayRow = 0
                    StartRow = CurRow + 1
                    Erase tmpData
                    ReDim xlData(Counter, 0)
                End If
                    UpdStatusBar picStatus, wiStatus
         
                rsTmp.MoveNext
                CurRow = CurRow + 1
            Loop
            rsTmp.Close
            Set rsTmp = Nothing
                        
            .SaveAs txtFilePath.Text
            xlApp.Quit
        End With

        UpdStatusBar picStatus, 100, True
        
    OutputPoData = True
    
    Set xlApp = Nothing
    
    wsSql = " DELETE FROM RPTEX001 WHERE RPTUSRID = '" & gsUserID & "'"
    wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
   cnCon.Execute wsSql
    
    Exit Function
    
OutputPoData_Err:
    
    MsgBox "OutputPoData Error!"
    
NoDataInRpt_Err:
    wsSql = " DELETE FROM RPTEX001 WHERE RPTUSRID = '" & gsUserID & "'"
    wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
   cnCon.Execute wsSql
    xlApp.Quit
    Set xlApp = Nothing
    rsTmp.Close
    Set rsTmp = Nothing
    UpdStatusBar picStatus, 0, True
    Exit Function
   
End Function



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
    If wcCombo.Enabled = True Then wcCombo.SetFocus

End Sub
