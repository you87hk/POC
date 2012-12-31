VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIM002 
   Caption         =   "Upload Stock In"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9270
   Icon            =   "frmIM002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   9270
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fmeFilePath 
      Caption         =   "Input Path"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   8730
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Index           =   0
         Left            =   8160
         TabIndex        =   9
         Top             =   525
         Width           =   288
      End
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   525
         Width           =   6312
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Input Sheet Path:"
         Height          =   645
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   1470
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4080
      Width           =   8565
   End
   Begin VB.Frame fmeFilePath 
      Caption         =   "Output Path"
      Height          =   1500
      Index           =   1
      Left            =   270
      TabIndex        =   3
      Top             =   2130
      Width           =   8655
      Begin VB.TextBox txtFilePath 
         Height          =   288
         Index           =   1
         Left            =   1800
         TabIndex        =   0
         Text            =   "ABCDEFGHIJKLMNOPQRS-"
         Top             =   600
         Width           =   6300
      End
      Begin VB.CommandButton cmdFilePath 
         Caption         =   "..."
         Height          =   288
         Index           =   1
         Left            =   8160
         TabIndex        =   1
         Top             =   600
         Width           =   288
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Export Sheet Path:"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   600
         Width           =   1470
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
            Picture         =   "frmIM002.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":1910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":1D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":207C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":2F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":33A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIM002.frx":3C82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
   Begin VB.Label lblStatus 
      Caption         =   "Progress:"
      Height          =   252
      Left            =   456
      TabIndex        =   2
      Top             =   3768
      Width           =   8376
   End
End
Attribute VB_Name = "frmIM002"
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
Private Const ItmCode = 2
Private Const WhsCode = 3
Private Const ItmNam = 4
Private Const ItPub = 5
Private Const Qty = 6
Private Const Uom = 7

Private Const ELine = 1
Private Const EDocNo = 2
Private Const EItmCode = 3
Private Const EItmNam = 4
Private Const EErrMsg = 5

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
    
    If LoadExcelData = True Then
    If PrintErrorSheet = True Then
        gsMsg = "錯誤 Excel 資料表已製作完成!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Call LoadExcel(txtFilePath(OutPath).Text)
    Else
        
        Name txtFilePath(inPath).Text As Left(txtFilePath(inPath).Text, Len(Trim(txtFilePath(inPath).Text)) - 4) & "_BK" & Right(Trim(txtFilePath(inPath).Text), 4)
        
        If UpdateStockIn Then
        gsMsg = "收貨已更新!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        End If
    End If
    End If
    
    Call Ini_Scr
    MousePointer = vbDefault
    
    Exit Sub
    
cmdOK_Err:

    gsMsg = "cmdOK Err!"
    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    MousePointer = vbDefault

End Sub





Private Sub cmdFilePath_Click(Index As Integer)

      On Error Resume Next

    If Trim(txtFilePath(Index).Text) <> "" Then
       ' If Dir(txtFilePath(Index).Text) <> "" Then
            'Dialog.dirDirectory.Path = txtFilePath.Text
            cdFilePath.InitDir = wsExcelPath
       ' End If
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
    
    wsFormID = "IM002"
End Sub

Private Sub Ini_Scr()


    
   Me.Caption = wsFormCaption

   
   If InStr(gsExcPath, ":\") Or InStr(gsExcPath, "\\") Then
        wsExcelPath = gsExcPath
    Else
        wsExcelPath = App.Path & "\" & gsExcPath
    End If
    
    txtFilePath(inPath).Text = wsExcelPath & "StockIn.xls"
   
    txtFilePath(OutPath).Text = wsExcelPath & "ERR_" & _
                    Format(Now, "YYYYMMDDHHMM") & ".xls"
   
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
   
   Set frmIM002 = Nothing

End Sub

Private Sub Ini_Caption()

   Call Get_Scr_Item("IM002", waScrItm)
   Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
  ' wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
   
  ' fmeFilePath(inPath).Caption = Get_Caption(waScrItm, "INPATH")
  ' fmeFilePath(OutPath).Caption = Get_Caption(waScrItm, "OUTPATH")
   
  ' lblFilePath.Caption = Get_Caption(waScrItm, "OUTPUTPATH")
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
   
    fmeFilePath(inPath).Caption = Get_Caption(waScrItm, "FMEINPATH")
    fmeFilePath(OutPath).Caption = Get_Caption(waScrItm, "FMEOUTPATH")
   
    lblFilePath(inPath).Caption = Get_Caption(waScrItm, "INPATH")
    lblFilePath(OutPath).Caption = Get_Caption(waScrItm, "OUTPATH")
   
    lblStatus.Caption = Get_Caption(waScrItm, "STATUS")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
   
   
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
    
    If Not Trim(txtFilePath(inIndex).Text) Like "*" & Dir(txtFilePath(inIndex)) & "*" Then
        gsMsg = "找不到檔案!"
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
    
    If Trim(txtFilePath(inIndex).Text) Like "*" & Dir(txtFilePath(inIndex)) & "*" And _
        Dir(txtFilePath(inIndex)) <> "" Then
        gsMsg = "檔案已存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtFilePath(inIndex).SetFocus
        Exit Function
    End If

End Select

    chk_txtFilePath = True

End Function

Private Function InputValidation() As Boolean

    InputValidation = False
    
   
    If chk_txtFilePath(inPath) = False Then
        txtFilePath(inPath).SetFocus
        Exit Function
    End If
    
    If chk_txtFilePath(OutPath) = False Then
        txtFilePath(OutPath).SetFocus
        Exit Function
    End If
    
    InputValidation = True

End Function


Private Function PrintErrorSheet() As Boolean
    
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
    Dim FileName As String
    Dim dTmpValue As Double
    
    On Error GoTo PrintErrorSheet_Err

    PrintErrorSheet = False
    
    
    
        wsSql = " SELECT RPTLINE, RPTDOCNO, RPTITMCODE, RPTITMNAM, RPTERRMSG "
        wsSql = wsSql & " FROM RPTERROR "
        wsSql = wsSql & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
        wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
        wsSql = wsSql & " ORDER BY RPTLINE, RPTITMCODE "
        rsTmp.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
        RecordNo = rsTmp.RecordCount
    
        If RecordNo <= 0 Then
           rsTmp.Close
           Set rsTmp = Nothing
           Exit Function
        End If
    
    wiStatus = 0
    Counter = EErrMsg
    
    'PRINT HEADING
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
         
        CurRow = 3
        StartRow = 3
        ArrayRow = 0
        xlApp.Workbooks.Open wsExcelPath & "DOCERROR.XLS", ReadOnly:=True
        
  '      FileName = ReadRs(rsRcd, "RPTPERIOD")
    
        
        lblStatus.Caption = "Generating File: " & txtFilePath(OutPath)

        
        With xlApp.Sheets("Sheet1")
          '  .Cells(1, 1).Value = "'P/O DATA UPLOAD FOR PERIOD" & FileName
            ' insert details
            ReDim xlData(Counter, 0)
            
            Do While Not rsTmp.EOF
        
                wiStatus = wiStatus + Fix((1 / RecordNo) * (100))
                
                ArrayRow = ArrayRow + 1
                ReDim Preserve xlData(Counter, ArrayRow - 1)
        
                For CurCol = ELine To EErrMsg
                    Select Case CurCol
                    
                    Case ELine
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
                        
            .SaveAs txtFilePath(OutPath).Text
            xlApp.Quit
        End With

        UpdStatusBar picStatus, 100, True
        
    PrintErrorSheet = True
    
    Set xlApp = Nothing
    
    wsSql = " DELETE FROM RPTERROR WHERE RPTUSRID = '" & gsUserID & "'"
    wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
    cnCon.Execute wsSql
    
    Exit Function
    
PrintErrorSheet_Err:
    
    MsgBox "PrintErrorSheet Error!"
    wsSql = " DELETE FROM RPTERROR WHERE RPTUSRID = '" & gsUserID & "'"
    wsSql = wsSql & " AND RPTDTETIM = '" & wsGenDte & "' "
    cnCon.Execute wsSql
    xlApp.Quit
    Set xlApp = Nothing
    rsTmp.Close
    Set rsTmp = Nothing
    UpdStatusBar picStatus, 0, True
    Exit Function
   
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


Private Function LoadExcelData() As Boolean

    Dim xlApp As Excel.Application
    Dim wiRow As Long
    Dim wsSql As String
    Dim wiCtr As Integer
    Dim wsDte As String
    
    On Error GoTo LoadExcelData_Err
    
    LoadExcelData = False
    
   
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
   
    UpdStatusBar picStatus, 20
    
    'get format A data
    xlApp.Workbooks.Open (Trim$(txtFilePath(inPath))), ReadOnly:=True
    wsGenDte = Change_SQLDate(Now)
        
    UpdStatusBar picStatus, 40
    
    wiRow = 3
    lblStatus.Caption = "Loading Record: " & " " & wiRow - 3
    
    With xlApp.Worksheets(1)
        Do While Trim(.Cells(wiRow, DocNo).Value) <> ""
            wsSql = " EXEC USP_RPTIM002 '" & gsUserID & "', '" & wsGenDte & "', "
            wsSql = wsSql & wiRow & ", "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, DocNo).Value) & "', "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, ItmCode).Value) & "', "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, WhsCode).Value) & "', "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, ItmNam).Value) & "', "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, ItPub).Value) & "', "
            wsSql = wsSql & To_Value(.Cells(wiRow, Qty).Value) & ", "
            wsSql = wsSql & "'" & Set_Quote(.Cells(wiRow, Uom).Value) & "', "
            wsSql = wsSql & "'" & wsFormID & "', "
            wsSql = wsSql & gsLangID
            
            'For wiCtr = 1 To ANoOfCol
            '    wsSql = wsSql & ", '" & Set_Quote(.Cells(wiRow, ACol1 + (wiCtr - 1))) & "'"
            'Next
            
            cnCon.Execute wsSql
            'lblStatus.Caption = Get_Caption(waScrItm, "FORMATA") & " " & wiRow - 2
            lblStatus.Caption = "Loading Record: " & " " & wiRow - 2
            DoEvents
            
            wiRow = wiRow + 1
        Loop
    End With
    xlApp.Workbooks.Close
    UpdStatusBar picStatus, 80

       
    xlApp.Quit
    Set xlApp = Nothing
    UpdStatusBar picStatus, 100, True

    
    If wiRow = 3 Then
        gsMsg = "Excel 表內沒有資料 : " & txtFilePath(inPath).Text
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    LoadExcelData = True
    
    Exit Function
     
LoadExcelData_Err:
   ' MsgBox "LoasExcelData_Err!"
    MsgBox Err.Description
    On Error Resume Next
    xlApp.Quit
    Set xlApp = Nothing
    Exit Function
   
End Function


Private Function UpdateStockIn() As Boolean
    Dim adComand As New ADODB.Command
    
    On Error GoTo UpdateStockIn_Err
    
    UpdateStockIn = False
    
    cnCon.BeginTrans
    Set adComand.ActiveConnection = cnCon
        
    adComand.CommandText = "USP_UPDSTKIN"
    adComand.CommandType = adCmdStoredProc
    adComand.Parameters.Refresh
      
    Call SetSPPara(adComand, 1, gsUserID)
    Call SetSPPara(adComand, 2, wsGenDte)
    
    adComand.Execute
    
    cnCon.CommitTrans
    
    
    Set adComand = Nothing
    UpdateStockIn = True
    
    MousePointer = vbDefault
    
    Exit Function
    
UpdateStockIn_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adComand = Nothing
    
End Function

