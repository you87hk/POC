VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGLP003 
   Caption         =   "Upload Stock Balance "
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9270
   Icon            =   "frmGLP003.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9270
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fmeFilePath 
      Caption         =   "Input Path"
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   8730
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Index           =   0
         Left            =   8160
         TabIndex        =   2
         Top             =   525
         Width           =   288
      End
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   480
         Width           =   5715
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Input Sheet Path:"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4080
      Width           =   8565
   End
   Begin VB.Frame fmeFilePath 
      Caption         =   "Output Path"
      Height          =   1500
      Index           =   1
      Left            =   270
      TabIndex        =   7
      Top             =   2130
      Width           =   8655
      Begin VB.CheckBox chkPrint 
         Caption         =   "New Page with Each Customer:"
         Height          =   180
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   600
         Width           =   5700
      End
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Index           =   1
         Left            =   8160
         TabIndex        =   5
         Top             =   600
         Width           =   288
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Export Sheet Path:"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   8
         Top             =   600
         Width           =   2190
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
            Picture         =   "frmGLP003.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":1910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":1D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":207C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":2F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":33A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP003.frx":3C82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   10
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
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   210
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   2010
   End
   Begin VB.Label lblStatus 
      Caption         =   "Progress:"
      Height          =   252
      Left            =   456
      TabIndex        =   6
      Top             =   3768
      Width           =   8376
   End
End
Attribute VB_Name = "frmGLP003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim xlsApp As Object
Dim xlsWrkBook As Object
Dim xlsSheet As Object

Dim waAccItm As New XArrayDB

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

Private Const bsCol = 0
Private Const bsRow = 1
Private Const bsCell = 2
Private Const bsAccCode = 3

Private Const tcGo = "Go"
Private Const tcExit = "Exit"
Private Const tcCancel = "Cancel"


Private Sub cmdCancel()
    MousePointer = vbDefault
    Ini_Scr
    txtFilePath(inPath).SetFocus
End Sub

Private Sub cmdOK()

    MousePointer = vbHourglass
    
    On Error GoTo cmdOK_Err
    
    If InputValidation = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Init_Excel
    DoEvents
    Call LoadExcelData
    Call Update_Excel
    
     If chkPrint.Value = 1 Then
        xlsSheet.PrintOut
     End If
    
     xlsApp.ActiveWorkbook.SaveCopyAs (txtFilePath(OutPath))
    
    gsMsg = "Balance Sheet Has been produced!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    
    Call Quit_Excel
    
    Call Ini_Scr
    MousePointer = vbDefault
    
    Exit Sub
    
cmdOK_Err:

    gsMsg = "cmdOK Err!"
    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    MousePointer = vbDefault

End Sub

Private Sub Init_Excel()

    Dim Cell As Range
    
    Set xlsApp = CreateObject("Excel.Application")
    
        
    Set xlsWrkBook = xlsApp.Workbooks
    
    With xlsApp
        .Visible = True
        .Caption = "Printing Financial Statements"
        .WindowState = xlMinimized
    End With
    
    xlsWrkBook.Open (Trim$(txtFilePath(inPath)))
    
    Set xlsSheet = xlsApp.Worksheets("Template")
        
    
    
    waAccItm.ReDim 0, -1, bsCol, bsAccCode

      
    
End Sub

Private Sub LoadExcelData()
    
    Dim Cell As Range
    Dim wsPrtRng As String
    Dim wsAccCode As String
        
    MousePointer = vbHourglass
    
  
    
    With xlsSheet
        .Select
    
        wsPrtRng = .PageSetup.PrintArea
        If wsPrtRng = "" Then
            gsMsg = "Print Area Not Set!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            Exit Sub
        End If
    
        For Each Cell In .Range(wsPrtRng)
            If Left$(Trim$(Trim$(Cell.Text)), 1) = "_" Then
           
                wsAccCode = .Range("A" & Cell.Row).Text
                DoEvents
                
                waAccItm.AppendRows
                waAccItm(waAccItm.UpperBound(1), bsCol) = Cell.Column
                waAccItm(waAccItm.UpperBound(1), bsRow) = Cell.Row
                waAccItm(waAccItm.UpperBound(1), bsCell) = Mid$(Trim$(Cell.Text), 2)
                waAccItm(waAccItm.UpperBound(1), bsAccCode) = wsAccCode
                
            End If
        Next
    End With
    
    MousePointer = vbDefault

End Sub

Private Sub Update_Excel()

    Dim wsColNo As String
    Dim wsRowNo As String
    Dim wsParam As String
    Dim wsAccCode As String
    Dim wsTmpDte As String
    Dim wsTmpPrd As String
    Dim wsSQL As String
    Dim wsOutput As String
    
    Dim Cell As Range
    Dim wiCtr As Integer
    Dim wsAcNam As String
    Dim wsComNam As String
    Dim wsCurrEarn As String
    Dim wsRetainAC As String
    Dim wsRetainCode As String
    
    Dim adcmdSave As New ADODB.Command
    
    
    
    On Error GoTo Update_Error
    
    

    If waAccItm.UpperBound(1) < 0 Then
        gsMsg = "No Item to Print!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inPath).SetFocus
        Exit Sub
    End If
     
    MousePointer = vbHourglass
    
    wsTmpDte = medPrdFr & "/01"
    
    wsTmpPrd = Get_FiscalPeriod(wsTmpDte)
    wsCurrEarn = Get_CompanyFlag("CMPCURREARN")
 '   wsRetainAC = Get_CompanyFlag("CMPRETAINAC")
 '   wsRetainCode = Get_TableInfo("MSTCOA", "COAACCID = " & To_Value(wsRetainAC), "COAACCCODE")
    
    
    
    For wiCtr = 0 To waAccItm.UpperBound(1)
        wsColNo = waAccItm(wiCtr, bsCol)
        wsRowNo = waAccItm(wiCtr, bsRow)
        wsParam = waAccItm(wiCtr, bsCell)
        wsAccCode = waAccItm(wiCtr, bsAccCode)
        
        Select Case wsParam
            Case "COMP"     ' Show Company Name
               If gsLangID = "1" Then
               wsComNam = Get_TableInfo("MstCompany", "CmpID = 1", "CmpEngName")
               Else
               wsComNam = Get_TableInfo("MstCompany", "CmpID = 1", "CmpChiName")
               End If
               xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsComNam
                
            Case "ACNO"     ' Show Account No
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = "'" & wsAccCode
            
            Case "ACNAME"   ' Show Account Name
                wsAcNam = Get_TableInfo("MstCOA", "COAAccCode = '" & Set_Quote(wsAccCode) & "'", IIf(gsLangID = "2", "COACDESC", "COADESC"))
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = "'" & wsAcNam
            
            Case "ASAT"
            
                wsTmpDte = Left(wsTmpPrd, 4) & "/" & Right(wsTmpPrd, 2) & "/01"
                wsTmpDte = DateAdd("d", -1, DateAdd("m", 1, CDate(wsTmpDte)))
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsTmpDte
                
            Case "STRDTE"
            
                wsTmpDte = Left(wsTmpPrd, 4) & "/" & Right(wsTmpPrd, 2) & "/01"
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsTmpDte
                
            Case "STRYR"
            
                wsTmpDte = Left(wsTmpPrd, 4) & "/01/01"
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsTmpDte
                
            Case Else
                 wsParam = UCase(wsParam)
                
                Set adcmdSave.ActiveConnection = cnCon
                
                If UCase(Trim(wsCurrEarn)) = UCase(Trim(wsAccCode)) Then
                
                adcmdSave.CommandText = "USP_RPTGLP003_CURREARN"
                adcmdSave.CommandType = adCmdStoredProc
                adcmdSave.Parameters.Refresh
                
                Call SetSPPara(adcmdSave, 1, Left(medPrdFr, 4) & Right(medPrdFr, 2))
                Call SetSPPara(adcmdSave, 2, wsAccCode)
                Call SetSPPara(adcmdSave, 3, wsParam)
                adcmdSave.Execute
    
                wsOutput = GetSPPara(adcmdSave, 4)
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsOutput
                
                Else
                
                adcmdSave.CommandText = "USP_RPTGLP003"
                adcmdSave.CommandType = adCmdStoredProc
                adcmdSave.Parameters.Refresh
                
                Call SetSPPara(adcmdSave, 1, Left(medPrdFr, 4) & Right(medPrdFr, 2))
                Call SetSPPara(adcmdSave, 2, wsAccCode)
                Call SetSPPara(adcmdSave, 3, wsParam)
                adcmdSave.Execute
    
                wsOutput = GetSPPara(adcmdSave, 4)
                xlsSheet.Cells(Val(wsRowNo), Val(wsColNo)).Value = wsOutput
                
                If To_Value(wsOutput) = 0 Then
                xlsSheet.Rows(Val(wsRowNo)).RowHeight = 0
                End If
                
                End If
                
        End Select
    Next
    
    MousePointer = vbDefault
    
    Exit Sub
    
Update_Error:
    
    MsgBox "Update_Error!"
    MousePointer = vbDefault
    Set adcmdSave = Nothing
    
    Exit Sub
                        
    
End Sub


Private Sub cmdFilePath_Click(Index As Integer)

      On Error Resume Next

    If Trim(txtFilePath(Index).Text) <> "" Then
        If Dir(txtFilePath(Index).Text) <> "" Then
            'Dialog.dirDirectory.Path = txtFilePath.Text
            cdFilePath.InitDir = wsExcelPath
        End If
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
        txtFilePath(Index).Text = .FileName
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
    
    wsFormID = "GLP003"
End Sub

Private Sub Ini_Scr()
Dim wsCtlMth As String

    
   Me.Caption = wsFormCaption

   
   If InStr(gsExcPath, ":\") Or InStr(gsExcPath, "\\") Then
        wsExcelPath = gsExcPath
    Else
        wsExcelPath = App.Path & "\" & gsExcPath
    End If
    
    txtFilePath(inPath).Text = wsExcelPath & "BSTemplete.xls"
   
    txtFilePath(OutPath).Text = wsExcelPath & "BS_" & _
                    Format(Now, "YYYYMMDDHHMM") & ".xls"
                    
    Call SetPeriodMask(medPrdFr)
    
    wsCtlMth = getCtrlMth("GL")
    
    medPrdFr.Text = Left(wsCtlMth, 4) & "/" & Right(wsCtlMth, 2)
    
    chkPrint.Value = 1
   
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
   Set waAccItm = Nothing
   Set frmGLP003 = Nothing

End Sub

Private Sub Ini_Caption()

    Call Get_Scr_Item("GLP003", waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
   
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    fmeFilePath(inPath).Caption = Get_Caption(waScrItm, "FMEINPATH")
    fmeFilePath(OutPath).Caption = Get_Caption(waScrItm, "FMEOUTPATH")
   
    lblFilePath(inPath).Caption = Get_Caption(waScrItm, "INPATH")
    lblFilePath(OutPath).Caption = Get_Caption(waScrItm, "OUTPATH")
   
    lblStatus.Caption = Get_Caption(waScrItm, "STATUS")
    chkPrint.Caption = Get_Caption(waScrItm, "CHKPRINT")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

End Sub



Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub


Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        txtFilePath(0).SetFocus
        
    End If
End Sub

Private Sub medPrdFr_LostFocus()
        FocusMe medPrdFr, True
End Sub

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    If chk_RetPrd(medPrdFr) = False Then
        gsMsg = "Period Must > Minimin Date!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    End If
    
    chk_medPrdFr = True
End Function

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



Private Sub txtFilePath_DblClick(Index As Integer)

        Call LoadExcel(txtFilePath(Index).Text)

End Sub



Private Sub txtFilePath_GotFocus(Index As Integer)

    FocusMe txtFilePath(Index)
    
End Sub



Private Function chk_txtFilePath(inIndex As Integer) As Boolean

    chk_txtFilePath = False
    
 Select Case inIndex
 Case inPath
    
    If Trim(txtFilePath(inIndex)) = "" Then
        gsMsg = "請輸入路徑!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If
    
    If Not Trim(UCase(txtFilePath(inIndex).Text)) Like "*" & UCase(Dir(txtFilePath(inIndex))) & "*" Then
        gsMsg = "找不到檔案!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If

    If UCase(Right(txtFilePath(inIndex).Text, 4)) <> ".XLS" Then
        gsMsg = "Input Must a Excel File!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If

Case OutPath


    
    If Trim(txtFilePath(inIndex)) = "" Then
        gsMsg = "請輸入路徑!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If
    
    If Trim(UCase(txtFilePath(inIndex).Text)) Like "*" & UCase(Dir(txtFilePath(inIndex))) & "*" And _
        Dir(txtFilePath(inIndex)) <> "" Then
        gsMsg = "檔案已存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If
    
    If Trim(txtFilePath(0)) = Trim(txtFilePath(1)) Then
        gsMsg = "Output file must not the same as Input !"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If

    If UCase(Right(txtFilePath(inIndex).Text, 4)) <> ".XLS" Then
        gsMsg = "Output Must a Excel File!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If


End Select

    chk_txtFilePath = True

End Function

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_medPrdFr = False Then
        Exit Function
    End If
    
   
    If chk_txtFilePath(inPath) = False Then
        Exit Function
    End If
    
    If chk_txtFilePath(OutPath) = False Then
        Exit Function
    End If
    
    InputValidation = True

End Function






Private Sub txtFilePath_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_txtFilePath(Index) = False Then
            Exit Sub
        End If
        If Index = inPath Then
        txtFilePath(OutPath).SetFocus
        Else
        txtFilePath(inPath).SetFocus
        End If

    End If
    
End Sub

Private Sub txtFilePath_LostFocus(Index As Integer)
    FocusMe txtFilePath(Index), True
End Sub



Private Function chk_RetPrd(inMedDte As Date) As Boolean
    Dim wsStrDte As String
    Dim wsEndDte As String
    Dim wsCtlDte As String
    Dim wiRetValue As Integer
    Dim wdMinDte As Date
    chk_RetPrd = False

    wiRetValue = Get_TableInfo("sysMonCtl", "MCMODNO = 'GL'", "MCKeepMn")
    
    wsCtlDte = getCtrlMth("GL", wsStrDte, wsEndDte)
    wdMinDte = DateAdd("M", To_Value(wiRetValue) * -1, (CDate(wsStrDte)))
    If CDate(inMedDte) < wdMinDte Then Exit Function
    
    chk_RetPrd = True
End Function


Private Sub Quit_Excel()

    On Error Resume Next
    
    With xlsApp
        .ActiveWorkbook.Saved = True
        .Quit
    End With
    
    Set xlsApp = Nothing
    Set xlsWrkBook = Nothing
    Set xlsSheet = Nothing
    
End Sub
