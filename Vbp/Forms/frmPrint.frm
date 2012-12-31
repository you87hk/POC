VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Printing"
   ClientHeight    =   7560
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9750
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9750
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame lblSavePath 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   8775
      Begin VB.CommandButton btnSavePath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6915
         Style           =   1  '圖片外觀
         TabIndex        =   23
         Tag             =   "K"
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDspSavePath 
         BorderStyle     =   1  '單線固定
         Caption         =   "\\SPSRV0\STEPPRO\REPORTPATH"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   405
         Width           =   6390
      End
      Begin VB.Label lblNoOfRecords 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "Records"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   2625
      End
   End
   Begin MSComDlg.CommonDialog cdPrinter 
      Left            =   9240
      Top             =   1320
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   792
      Top             =   360
   End
   Begin MSComDlg.CommonDialog cdSaveAs 
      Left            =   9240
      Top             =   960
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
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
      Left            =   264
      ScaleHeight     =   360
      ScaleWidth      =   9105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   405
      Width           =   9168
   End
   Begin TabDlg.SSTab tabFieldSelect 
      Height          =   4620
      Left            =   285
      TabIndex        =   17
      Top             =   2760
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8149
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Field Selection"
      TabPicture(0)   =   "frmPrint.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstSelect(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstSelect(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelect(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSelect(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSelect(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSelect(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSelect(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSelect(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkOnScreen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Sorting"
      TabPicture(1)   =   "frmPrint.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTopN"
      Tab(1).Control(1)=   "cmdSort(5)"
      Tab(1).Control(2)=   "cmdSort(4)"
      Tab(1).Control(3)=   "cmdSort(3)"
      Tab(1).Control(4)=   "cmdSort(2)"
      Tab(1).Control(5)=   "cmdSort(1)"
      Tab(1).Control(6)=   "cmdSort(0)"
      Tab(1).Control(7)=   "lstSort(0)"
      Tab(1).Control(8)=   "lstSort(1)"
      Tab(1).Control(9)=   "lblTopN"
      Tab(1).ControlCount=   10
      Begin VB.CheckBox chkOnScreen 
         Caption         =   "View on Screen"
         Height          =   324
         Left            =   312
         TabIndex        =   20
         Top             =   3864
         Value           =   1  '核取
         Width           =   3348
      End
      Begin VB.TextBox txtTopN 
         Height          =   288
         Left            =   -73320
         TabIndex        =   19
         Top             =   3864
         Width           =   1020
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   345
         Index           =   5
         Left            =   8475
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   2256
         Width           =   585
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   345
         Index           =   4
         Left            =   8475
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   1800
         Width           =   585
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "<<"
         Height          =   540
         Index           =   3
         Left            =   3792
         TabIndex        =   4
         Top             =   3168
         Width           =   1044
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "<"
         Height          =   540
         Index           =   2
         Left            =   3792
         TabIndex        =   3
         Top             =   2520
         Width           =   1044
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   ">>"
         Height          =   540
         Index           =   1
         Left            =   3792
         TabIndex        =   2
         Top             =   1200
         Width           =   1044
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   ">"
         Height          =   540
         Index           =   0
         Left            =   3792
         TabIndex        =   1
         Top             =   552
         Width           =   1044
      End
      Begin VB.CommandButton cmdSort 
         Height          =   345
         Index           =   5
         Left            =   -66525
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   2256
         Width           =   585
      End
      Begin VB.CommandButton cmdSort 
         Height          =   345
         Index           =   4
         Left            =   -66525
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   1800
         Width           =   585
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "<<"
         Height          =   540
         Index           =   3
         Left            =   -71208
         TabIndex        =   12
         Top             =   3168
         Width           =   1044
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "<"
         Height          =   540
         Index           =   2
         Left            =   -71208
         TabIndex        =   11
         Top             =   2520
         Width           =   1044
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   ">>"
         Height          =   540
         Index           =   1
         Left            =   -71208
         TabIndex        =   10
         Top             =   1200
         Width           =   1044
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   ">"
         Height          =   540
         Index           =   0
         Left            =   -71208
         TabIndex        =   9
         Top             =   552
         Width           =   1044
      End
      Begin MSComctlLib.ListView lstSelect 
         Height          =   3468
         Index           =   0
         Left            =   288
         TabIndex        =   0
         Top             =   288
         Width           =   3348
         _ExtentX        =   5900
         _ExtentY        =   6112
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstSelect 
         Height          =   3468
         Index           =   1
         Left            =   4992
         TabIndex        =   5
         Top             =   288
         Width           =   3348
         _ExtentX        =   5900
         _ExtentY        =   6112
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstSort 
         Height          =   3468
         Index           =   0
         Left            =   -74712
         TabIndex        =   8
         Top             =   288
         Width           =   3348
         _ExtentX        =   5900
         _ExtentY        =   6112
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstSort 
         Height          =   3468
         Index           =   1
         Left            =   -70008
         TabIndex        =   13
         Top             =   288
         Width           =   3348
         _ExtentX        =   5900
         _ExtentY        =   6112
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblTopN 
         Caption         =   "Top N Records:"
         Height          =   252
         Left            =   -74616
         TabIndex        =   18
         Top             =   3864
         Width           =   1164
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":10F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":141A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":187A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Preview (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Excel (F5)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse"
            Object.ToolTipText     =   "Browse (F6)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Printer"
            Object.ToolTipText     =   "Printer Setup (F9)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Detail"
            Object.ToolTipText     =   "Detail (F10)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Declare the constant value
'Option


Private Const cmdPreview   As Integer = 0
Private Const cmdPrinter   As Integer = 1
Private Const cmdExcel     As Integer = 2
Private Const cmdListView  As Integer = 3
'Command button for printing
Private Const PrintOK      As Integer = 0
Private Const PrintCancel  As Integer = 1
Private Const PrintSave    As Integer = 2
Private Const PrintPrinter As Integer = 3
Private Const PrintDetail  As Integer = 4
Private Const PrintFont    As Integer = 5
'Tab control
Private Const TabSelect    As Integer = 0
Private Const TabSort      As Integer = 1
'Command button for listview
Private Const MoveRight    As Integer = 0
Private Const MoveRightAll As Integer = 1
Private Const MoveLeft     As Integer = 2
Private Const MoveLeftAll  As Integer = 3
Private Const MoveUp       As Integer = 4
Private Const MoveDown     As Integer = 5
'Listview for from or to
Private Const PrintFrom    As Integer = 0
Private Const PrintTo      As Integer = 1
'Listview field subitem
Private Const ItemField    As Integer = 1
Private Const ItemNumFlag  As Integer = 2

Private Const cmdStart     As Integer = 1
Private Const cmdProcess   As Integer = 2
Private Const cmdFail      As Integer = 3

Private Const SummaryHeight   As Long = 3000
Private Const DetailHeight    As Long = 8000
Private Const FormWidth       As Long = 9816


Private Const tcPreview = "Preview"
Private Const tcPrint = "Print"
Private Const tcExcel = "Excel"
Private Const tcBrowse = "Browse"
Private Const tcPrinter = "Printer"
Private Const tcDetail = "Detail"
Private Const tcExit = "Exit"

'variable for new property
Private msTableID      As String
Private msReportID   As String
Private msStoreP     As String
Private msRptDteTim  As String
Private msRptTitle   As String
Private msRptName   As String
Private msSelection  As Variant
'misc. variable
Dim wsServer      As String
Dim wsDatabase    As String
Dim wsUser        As String
Dim wsPassword    As String
Dim wsSavePath    As String
Dim wiStatus      As Integer
Dim wiNoOfCopy    As Long
Dim wsQuery       As String
Dim wiNoOfRecords As Long
Dim wsAction      As Integer
Dim waScrItm      As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wiStartFlg    As Boolean
Dim wsRptPath As String
Dim wsExcPath As String

    
Private wsMsg As String

'Dim gdbSTEPPro As New ADODB.Connection
Dim WithEvents wdbCon As ADODB.Connection
Attribute wdbCon.VB_VarHelpID = -1


Property Get StoreP() As String

   StoreP = msStoreP
   
End Property

Property Let StoreP(ByVal NewStoreP As String)

   msStoreP = NewStoreP
   
End Property

Property Get TableID() As String

   TableID = msTableID
   
End Property

Property Let TableID(ByVal NewTableID As String)

   msTableID = NewTableID
   
End Property

Property Get RptTitle() As String

   RptTitle = msRptTitle
   
End Property
Property Let RptTitle(ByVal NewRptTitle As String)

   msRptTitle = NewRptTitle
   
End Property

Property Get RptName() As String

   RptName = msRptName
   
End Property
Property Let RptName(ByVal NewRptName As String)

   msRptName = NewRptName
   
End Property
Property Get RptDteTim() As String

   RptDteTim = msRptDteTim
   
End Property

Property Let RptDteTim(ByVal NewRptDteTim As String)

   msRptDteTim = NewRptDteTim
   
End Property

Property Get ReportID() As String

   ReportID = msReportID
   
End Property

Property Get Selection() As Variant

   Selection = msSelection
   
End Property

Property Let Selection(ByVal NewSelection As Variant)

   msSelection = NewSelection
   
End Property

Property Let ReportID(ByVal NewReportID As String)

   msReportID = NewReportID

End Property

Private Sub btnSavePath_Click()

      On Error Resume Next
      
      With cdSaveAs
         .DialogTitle = "Save Excel File to "
         .InitDir = gsExcPath
         .FileName = lblDspSavePath
         .Filter = "Excel File (*.xls)|*.xls"
         .CancelError = True
         .ShowSave
      
         If Err.Number <> cdlCancel Then
            lblDspSavePath = .FileName
         End If
      
      End With
      
      On Error GoTo 0
      
      
End Sub



Private Sub cmdSelect_Click(Index As Integer)

   Select Case Index
   Case MoveRight
      MoveSelectItem PrintFrom, False
   Case MoveRightAll
      MoveSelectItem PrintFrom, True
   Case MoveLeft
      MoveSelectItem PrintTo, False
   Case MoveLeftAll
      MoveSelectItem PrintTo, True
   Case MoveUp
      MovePosition MoveUp, TabSelect
   Case MoveDown
      MovePosition MoveDown, TabSelect
   End Select
   
End Sub

Private Sub cmdSort_Click(Index As Integer)

   Select Case Index
   Case MoveRight
      MoveSortItem PrintFrom, False
   Case MoveRightAll
      MoveSortItem PrintFrom, True
   Case MoveLeft
      MoveSortItem PrintTo, False
   Case MoveLeftAll
      MoveSortItem PrintTo, True
   Case MoveUp
      MovePosition MoveUp, TabSort
   Case MoveDown
      MovePosition MoveDown, TabSort
   End Select

End Sub

Private Sub Form_Activate()

   Me.MousePointer = vbHourglass

   Select Case wsAction
   Case cmdStart
      'Initialize Form Position
      Me.Height = SummaryHeight
      Me.Width = FormWidth
      Me.Top = (Screen.Height - DetailHeight) / 2
      Me.Left = (Screen.Width - Me.Width) / 2
      
   End Select
   
   Me.Refresh
   DoEvents
   If wiStartFlg = True Then
      wiStartFlg = False
      
      Ini_Caption
      
      Ini_Scr
      
      Set wdbCon = New ADODB.Connection
      
      With wdbCon
         .Provider = "SQLOLEDB"
         .ConnectionTimeout = 10
         .CursorLocation = adUseClient
         .ConnectionString = gsConnectString
         .Open
      End With
      
   End If
   
   Select Case wsAction
   Case cmdStart
      If RunStoredProcedure = False Then
         Me.MousePointer = vbDefault
         Unload Me
         Exit Sub
      End If
   'Case cmdFail
   '   Unload Me
   End Select
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        Case vbKeyF2
            
            If tbrProcess.Buttons(tcPreview).Enabled = True Then
                If gsRTAccess = "Y" Then
                Access_RunTimePrint cmdPreview
                Else
                Access_Print cmdPreview
                End If
            End If
        
        
        Case vbKeyF3
            If tbrProcess.Buttons(tcPrint).Enabled = True Then
                Access_Print cmdPrinter
                If gsRTAccess = "Y" Then
                Access_RunTimePrint cmdPrinter
                Else
                Access_Print cmdPrinter
                End If
            End If
            
        
        Case vbKeyF5
            If tbrProcess.Buttons(tcExcel).Enabled = True Then
                PrintExcel
            End If
       
        
        Case vbKeyF6
            If tbrProcess.Buttons(tcBrowse).Enabled = True Then
                PrintListView
            End If
        
        Case vbKeyF9
            If tbrProcess.Buttons(tcPrinter).Enabled = True Then
                PrinterSetup
            End If

            
        Case vbKeyF10
            
            PrintDetailForm
            
        Case vbKeyF12
            FormExit
            
    End Select
    
    
    KeyCode = vbDefault
    
       
End Sub

Private Sub Form_Load()

   Me.MousePointer = vbHourglass
   Me.Visible = False
   Me.KeyPreview = True
   
   wiStartFlg = True
   wsAction = cmdStart
   'Set wdbCon = gdbStepPro
 '  optPrint(cmdPreview).Value = True
   
   Me.MousePointer = vbDefault
   
End Sub


Private Sub Form_LostFocus()

   'Unload Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' SaveUserDefault
   
   wdbCon.Close
   Set wdbCon = Nothing

   cnCon.Execute "DELETE FROM RPT" & Me.TableID & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' AND RPTDTETIM = '" & Change_SQLDate(Me.RptDteTim) & "' "
   'gdbSTEPPro.Close
   'Set gdbSTEPPro = Nothing
   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set frmPrint = Nothing
   Me.MousePointer = vbDefault
   
End Sub


Private Function RunStoredProcedure() As Boolean

   RunStoredProcedure = False
   wsAction = cmdProcess

   If Trim(Me.StoreP) = "" Then
        wsMsg = "No Store Procedure !Please Check!"
        MsgBox wsMsg, vbOKOnly, gsTitle
      Exit Function
   End If
   
   Timer1.Enabled = True
   wdbCon.Execute Me.StoreP, , adAsyncExecute
   
   RunStoredProcedure = True

End Function

Private Sub lstSelect_BeforeLabelEdit(Index As Integer, Cancel As Integer)

   Cancel = True
   
End Sub

Private Sub lstSelect_DblClick(Index As Integer)

   MoveSelectItem Index, False
   
End Sub

Private Sub lstSelect_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)

   If Item.ListSubItems(ItemNumFlag).Text <> "N" Then
      Item.Checked = False
   End If
   
End Sub

Private Sub lstSort_BeforeLabelEdit(Index As Integer, Cancel As Integer)

   Cancel = True
   
End Sub

Private Sub lstSort_DblClick(Index As Integer)

   MoveSortItem Index, False
   
End Sub

Private Sub tabFieldSelect_Click(PreviousTab As Integer)

   Select Case tabFieldSelect.Tab
   Case TabSelect
      
   Case TabSort
      'UpdateListSort
      
   End Select
   
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)

 Select Case Button.Key
        Case tcPreview
            
                If gsRTAccess = "Y" Then
                Access_RunTimePrint cmdPreview
                Else
                Access_Print cmdPreview
                End If
        Case tcPrint
                If gsRTAccess = "Y" Then
                Access_RunTimePrint cmdPrinter
                Else
                Access_Print cmdPrinter
                End If

        Case tcExcel
            PrintExcel
        Case tcBrowse
            PrintListView
        Case tcPrinter
            PrinterSetup
        Case tcDetail
            PrintDetailForm
        Case tcExit
            FormExit
    End Select
    
      
End Sub

Private Sub Timer1_Timer()
   wiStatus = 100
   If wiStatus < 100 Then
      wiStatus = wiStatus + 1
      Timer1.Interval = Timer1.Interval + 200
   Else
      wiStatus = 0
      Timer1.Enabled = False
   End If
   UpdateStatus picStatus, wiStatus
    
End Sub


Private Sub txtTopN_GotFocus()

   FocusMe txtTopN
   
End Sub

Private Sub txtTopN_KeyPress(KeyAscii As Integer)

   Call Chk_InpNum(KeyAscii, txtTopN.Text, False, False)
   
End Sub



Private Sub wdbCon_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   
   On Error GoTo ExecuteComplete_Err

   Dim adReport As New ADODB.Recordset
   
   wsQuery = " SELECT * FROM RPT" & Me.TableID
   wsQuery = wsQuery & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
   wsQuery = wsQuery & "   AND RPTDTETIM = '" & Change_SQLDate(Me.RptDteTim) & "' "

   adReport.Open wsQuery, cnCon, adOpenStatic, adLockOptimistic
   
   If adReport.RecordCount = 0 Then
   'If RecordsAffected = 0 Then
      Timer1.Enabled = False
      UpdateStatus picStatus, 0
      wsAction = cmdFail
      MsgBox "No Report Data!Check the Date Range!"
      Exit Sub
      'Unload Me
   Else
      wiNoOfRecords = adReport.RecordCount
      lblNoOfRecords.Caption = lblNoOfRecords.Caption & wiNoOfRecords
      If To_Value(txtTopN.Text) = 0 Then
         txtTopN = adReport.RecordCount
      Else
         If To_Value(txtTopN.Text) > adReport.RecordCount Then
            txtTopN.Text = adReport.RecordCount
         End If
      End If
   End If
   
   'UpdateStatus picStatus, 0
  Call SetButtonStatus("All")
 '  Enable_PrintOption True
   If Dir(wsRptPath & gsDBName) = "" Then
      Call SetButtonStatus("NoPrint")
   Else
        Call SetButtonStatus("Print")
   End If
   
   LoadListView

   Timer1.Enabled = False
   For wiStatus = wiStatus To 99
      wiStatus = wiStatus + 1
      UpdateStatus picStatus, wiStatus
   Next
   UpdateStatus picStatus, 100, True
   adReport.Close
   Set adReport = Nothing
   
   Exit Sub
   
ExecuteComplete_Err:
   MsgBox "ExecuteComplete Module Error! " & Err.Number & ", " & Err.Description
   Timer1.Enabled = False
   UpdateStatus picStatus, 0
   On Error Resume Next
   'adReport.Close
   Set wdbCon = Nothing
   Set adReport = Nothing


End Sub

Private Sub Ini_Scr()
   
   Dim wiStrPos As Integer
   Dim wiSemiPos As Integer
   
   Me.Visible = True
   
  ' cmdSelect(MoveUp).Picture = LoadResPicture("UP-ARROW", vbResBitmap)
 '  cmdSelect(MoveDown).Picture = LoadResPicture("DOWN-ARROW", vbResBitmap)
 '  cmdSort(MoveUp).Picture = LoadResPicture("UP-ARROW", vbResBitmap)
 '  cmdSort(MoveDown).Picture = LoadResPicture("DOWN-ARROW", vbResBitmap)
   
   lstSelect(PrintTo).ToolTipText = "Select Desc"
   lstSort(PrintTo).ToolTipText = "Sort Desc"
  
   wiStrPos = InStr(1, gsConnectString, "Data Source=", vbTextCompare)
   wiSemiPos = InStr(wiStrPos, gsConnectString, ";", vbTextCompare)
   wsServer = Mid(gsConnectString, wiStrPos + 12, wiSemiPos - (wiStrPos + 12))
   
   wiStrPos = InStr(1, gsConnectString, "Initial Catalog=", vbTextCompare)
   wiSemiPos = InStr(wiStrPos, gsConnectString, ";", vbTextCompare)
   wsDatabase = Mid(gsConnectString, wiStrPos + 16, wiSemiPos - (wiStrPos + 16))
   
   wiStrPos = InStr(1, gsConnectString, "User ID=", vbTextCompare)
   wiSemiPos = InStr(wiStrPos, gsConnectString, ";", vbTextCompare)
   wsUser = Mid(gsConnectString, wiStrPos + 8, wiSemiPos - (wiStrPos + 8))
   
   wiStrPos = InStr(1, gsConnectString, "Password=", vbTextCompare)
   'wiSemiPos = InStr(wiStrPos, gsConnectionString, ";", vbTextCompare)
   wsPassword = Mid(gsConnectString, wiStrPos + 9)
     
   'Initialize Form Position
   'Me.Height = SummaryHeight
   'Me.Width = FormWidth
   'Me.Top = (Screen.Height - DetailHeight) / 2
   'Me.Left = (Screen.Width - Me.Width) / 2

   'Initialize enable of button and option box
   Call SetButtonStatus("None")
 '  Enable_PrintOption False
   
   tabFieldSelect.Tab = TabSelect

   'Initialize Printer dialog box
   With cdPrinter
      .flags = &H100000 Or &H4& Or &H80& Or &H100&
      .PrinterDefault = True
      .CancelError = True
      wiNoOfCopy = 1
      wiNoOfCopy = Printer.Copies
   End With
   
   'With cdFont
   '   .flags = cdlCFBoth Or cdlCFANSIOnly
   '   .CancelError = True
   'End With
   
   Ini_ListView
   
       ' Create new database in Microsoft Access window.
    If InStr(gsExcPath, ":\") Or InStr(gsExcPath, "\\") Then
        wsExcPath = gsExcPath
    Else
        wsExcPath = App.Path & "\" & gsExcPath
    End If
    
    
    If InStr(gsRptPath, ":\") Or InStr(gsRptPath, "\\") Then
        wsRptPath = gsRptPath
    Else
        wsRptPath = App.Path & "\" & gsRptPath
    End If
   

    lblDspSavePath = wsExcPath & Me.TableID & "_" & _
                    Format(Now, "YYYYMMDDHHMM") & ".XLS"

   
   
   txtTopN.Text = 0

End Sub

Private Sub PrintExcel()
   
   Dim xlApp      As Object
   Dim xlSheet1   As Object
   Dim adDetail   As New ADODB.Recordset
   Dim adSummary  As New ADODB.Recordset
      
   Dim wsSQL      As String   'SQL statement
   Dim NoOfFields As Integer  'Number of fields
   Dim NoOfSum    As Integer  'Number of summary fields
   Dim wsFields   As String   'Select field string
   Dim wsOrder    As String   'Order by field string
   Dim wsSum      As String   'Sum up field string
   Dim wsGroup    As String   'Grouping field string
   Dim wiCtr      As Integer
   Dim NoOfRecord As Long     'Number of Record
   Dim RowPack    As Integer  'Number of rows copy from array to excel -> depends on no. of col
   Dim wsText     As String
   Dim wsMid      As String
   Dim wiStatus   As Double
   Dim xlData     As Variant
   Dim tmpData    As Variant
   Dim StartRow   As Long
   Dim ArrayRow   As Long
   Dim CurRow     As Long
   Dim CurCol     As Integer
   Dim inpParent  As Variant
   Dim inpDate    As String
   Dim i          As Long
   Dim J          As Long
      
   On Error GoTo Excel_Print_Err1
   
   'if no field is selected then exit
   If lstSelect(PrintTo).ListItems.Count = 0 Then Exit Sub
   
   UpdateStatus picStatus, 5
   
   'CONSTRUCT THE SELECT FIELDS STRING AND SUMMARY STRING
   wsFields = ""
   wsSum = ""
   NoOfSum = 0
   With lstSelect(PrintTo)
      NoOfFields = .ListItems.Count
      For wiCtr = 1 To NoOfFields
         wsFields = wsFields & .ListItems(wiCtr).ListSubItems(ItemField).Text & ", "
         'NoOfFields = NoOfFields + 1
         If .ListItems(wiCtr).Checked = True Then
            If UCase(.ListItems(wiCtr).ListSubItems(ItemNumFlag).Text) = "N" Then
               wsSum = wsSum & "SUM(" & .ListItems(wiCtr).ListSubItems(ItemField).Text & "), "
               '"COUNT(DISTINCT " & .ListItems(wiCtr).ListSubItems(ItemField).Text & "), ")
               NoOfSum = NoOfSum + 1
            End If
         End If
      Next
   End With
   wsFields = Mid$(wsFields, 1, Len(Trim(wsFields)) - 1)
   If Len(Trim(wsSum)) > 0 Then
      wsSum = Mid$(wsSum, 1, Len(Trim(wsSum)) - 1)
   End If
   
   'CONSTRUCT THE SORTING STRING AND GROUPING STRING
   wsOrder = ""
   wsGroup = ""
   With lstSort(PrintTo)
      For wiCtr = 1 To .ListItems.Count
         wsOrder = wsOrder & .ListItems(wiCtr).ListSubItems(ItemField).Text & ", "
         If .ListItems(wiCtr).Checked = True Then
            wsGroup = wsGroup & .ListItems(wiCtr).ListSubItems(ItemField).Text & ", "
            NoOfSum = NoOfSum + 1
         End If
      Next
   End With
   If Len(Trim(wsOrder)) > 0 Then
      wsOrder = Mid$(wsOrder, 1, Len(Trim(wsOrder)) - 1)
      If Len(Trim(wsGroup)) > 0 Then
         wsGroup = Mid$(wsGroup, 1, Len(Trim(wsGroup)) - 1)
      End If
   End If
   
   'Construct the detail select statement
   If To_Value(txtTopN.Text) <> wiNoOfRecords Then
      wsSQL = " SELECT TOP " & To_Value(txtTopN.Text) & " WITH TIES "
      wsSQL = wsSQL & wsFields & " FROM RPT" & Me.TableID
   Else
      wsSQL = " SELECT " & wsFields & " FROM RPT" & Me.TableID
   End If
   wsSQL = wsSQL & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
   wsSQL = wsSQL & "   AND RPTDTETIM = '" & Change_SQLDate(Me.RptDteTim) & "' "
   If Trim(wsOrder) <> "" Then
      wsSQL = wsSQL & " ORDER BY " & wsOrder
   End If
   
   adDetail.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
   UpdateStatus picStatus, 10
   
   'Construct the summary select statement
   If Trim(wsGroup) <> "" And Trim(wsSum) <> "" And _
      To_Value(txtTopN.Text) = wiNoOfRecords Then
      wsSQL = " SELECT " & wsGroup & ", " & wsSum & " FROM RPT" & Me.TableID
      wsSQL = wsSQL & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
      wsSQL = wsSQL & "   AND RPTDTETIM = '" & Change_SQLDate(Me.RptDteTim) & "' "
      wsSQL = wsSQL & " GROUP BY " & wsGroup & " WITH CUBE "
      wsSQL = wsSQL & " ORDER BY " & wsGroup
      
      adSummary.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   End If
   
   UpdateStatus picStatus, 20
   
   If adDetail.RecordCount > 0 Then
      NoOfRecord = adDetail.RecordCount
   Else
      Exit Sub
   End If
   
   On Error GoTo Excel_Print_Err2
   
   Set xlApp = CreateObject("EXCEL.Application")
   xlApp.Visible = False
   xlApp.SheetsInNewWorkbook = 1
   xlApp.Workbooks.Add
   Set xlSheet1 = xlApp.Workbooks(1).Worksheets(1)
   
   If Init_Excel(xlSheet1, xlApp) = False Then GoTo Excel_Print_Err1
   
   UpdateStatus picStatus, 30
   
   'POPULATE EXCEL WORKSHEET HEADER
   With xlSheet1
      For wiCtr = 1 To NoOfFields
         .Cells(1, wiCtr).Value = lstSelect(PrintTo).ListItems(wiCtr).Text
      Next
   End With
   
   With xlSheet1.Range(xlSheet1.Cells(1, 1).Address, xlSheet1.Cells(1, NoOfFields).Address)
      .Interior.ColorIndex = 15
      .WrapText = True
      .Borders(xlTop).Weight = xlMedium
      .Borders(xlRight).Weight = xlMedium
      .Borders(xlLeft).Weight = xlMedium
      .Borders(xlBottom).Weight = xlMedium
      '.AutoFit
   End With
   
   UpdateStatus picStatus, 35
   
   ' INSERT DETAILS
   CurRow = 2
   StartRow = 2
   ArrayRow = 0
   wiStatus = 0
   RowPack = 100
   If NoOfFields > 20 Then
      RowPack = 50
   End If
   ReDim xlData(NoOfFields, 0)
    
   Do Until adDetail.EOF
     
      wiStatus = wiStatus + ((1 / NoOfRecord) * 40)
      ArrayRow = ArrayRow + 1
      ReDim Preserve xlData(NoOfFields, ArrayRow - 1)
     
      For CurCol = 1 To NoOfFields
        
         Select Case adDetail(CurCol - 1).Type
         Case adDate, adDBDate, adDBTime, adDBTimeStamp
            xlData(CurCol - 1, ArrayRow - 1) = Dsp_Date(ReadRs(adDetail, CurCol - 1), , True)
            
         Case adBinary, adCurrency, adDecimal, adDouble, adInteger, _
              adNumeric, adSingle, adSmallInt, adTinyInt
            xlData(CurCol - 1, ArrayRow - 1) = To_Value(ReadRs(adDetail, CurCol - 1))
            
         Case adLongVarBinary, adLongVarChar, adLongVarWChar
            inpParent = Trim(adDetail(CurCol - 1).GetChunk(2048))
            inpParent = IIf(IsNull(inpParent) = True, "", inpParent)
            wsText = ""
            For wiCtr = 1 To Len(inpParent)
               wsMid = Mid(inpParent, wiCtr, 1)
               If wsMid = Chr(13) Then
                   wsText = wsText & " "
               Else
                   wsText = wsText & wsMid
               End If
               xlData(CurCol - 1, ArrayRow - 1) = wsText
            Next
        
         Case Else    'string or unknown type
            xlData(CurCol - 1, ArrayRow - 1) = CStr("'" & "" & ReadRs(adDetail, CurCol - 1))
           
         End Select
      Next
     
      If ArrayRow = RowPack Or CurRow = NoOfRecord + 1 Then    '1 means the header row
         ReDim tmpData(UBound(xlData, 2), UBound(xlData, 1))
         For i = 0 To UBound(xlData, 1)
            For J = 0 To UBound(xlData, 2)
               tmpData(J, i) = xlData(i, J)
            Next
         Next
         xlSheet1.Range(xlSheet1.Cells(StartRow, 1).Address, xlSheet1.Cells(CurRow, NoOfFields).Address).Value = tmpData
   
         ArrayRow = 0
         StartRow = CurRow + 1
         Erase tmpData
         ReDim xlData(NoOfFields, 0)
     End If
     UpdateStatus picStatus, CInt(wiStatus) + 35
     adDetail.MoveNext
     CurRow = CurRow + 1
   Loop
   
   adDetail.Close
   Set adDetail = Nothing
   UpdateStatus picStatus, 75
     
   If Trim(wsGroup) = "" Or Trim(wsSum) = "" Or _
      To_Value(txtTopN.Text) <> wiNoOfRecords Then GoTo PrintExcel_Save
   
   'DISPLAY SUMMARY PART
   CurRow = CurRow + 4
   CurCol = 1
   With lstSort(PrintTo)
      For wiCtr = 1 To .ListItems.Count
         If .ListItems(wiCtr).Checked Then
            xlSheet1.Cells(CurRow, CurCol).Value = .ListItems(wiCtr).Text
            CurCol = CurCol + 1
         End If
      Next
   End With
   
   With lstSelect(PrintTo)
      For wiCtr = 1 To .ListItems.Count
         If .ListItems(wiCtr).Checked Then
            If UCase(.ListItems(wiCtr).ListSubItems(ItemNumFlag).Text) = "N" Then
               xlSheet1.Cells(CurRow, CurCol).Value = .ListItems(wiCtr).Text
               CurCol = CurCol + 1
            End If
         End If
      Next
   End With
   
   With xlSheet1.Range(xlSheet1.Cells(CurRow, 1).Address, xlSheet1.Cells(CurRow, NoOfSum).Address)
      .Interior.ColorIndex = 48
      .WrapText = True
      .Borders(xlTop).Weight = xlMedium
      .Borders(xlRight).Weight = xlMedium
      .Borders(xlLeft).Weight = xlMedium
      .Borders(xlBottom).Weight = xlMedium
      '.AutoFit
   End With
   
   UpdateStatus picStatus, 80
   
   ' INSERT DETAILS
   CurRow = CurRow + 1
   StartRow = CurRow
   ArrayRow = 0
   wiStatus = 0
   RowPack = 100
   If NoOfSum > 20 Then
      RowPack = 50
   End If
   ReDim xlData(NoOfSum, 0)
    
   Do Until adSummary.EOF
     
      wiStatus = wiStatus + ((1 / adSummary.RecordCount) * 15)
      ArrayRow = ArrayRow + 1
      ReDim Preserve xlData(NoOfSum, ArrayRow - 1)
     
      For CurCol = 1 To NoOfSum
        
         Select Case adSummary(CurCol - 1).Type
         Case adDate, adDBDate, adDBTime, adDBTimeStamp
            xlData(CurCol - 1, ArrayRow - 1) = Dsp_Date(ReadRs(adSummary, CurCol - 1), , True)
            
         Case adBinary, adCurrency, adDecimal, adDouble, adInteger, _
              adNumeric, adSingle, adSmallInt, adTinyInt
            xlData(CurCol - 1, ArrayRow - 1) = To_Value(ReadRs(adSummary, CurCol - 1))
            
         Case adLongVarBinary, adLongVarChar, adLongVarWChar
            inpParent = Trim(adSummary(CurCol - 1).GetChunk(2048))
            wsText = ""
            For wiCtr = 1 To Len(inpParent)
               wsMid = Mid(inpParent, wiCtr, 1)
               If wsMid = Chr(13) Then
                   wsText = wsText & " "
               Else
                   wsText = wsText & wsMid
               End If
            Next
            xlData(CurCol - 1, ArrayRow - 1) = wsText
        
         Case Else    'string or unknown type
            xlData(CurCol - 1, ArrayRow - 1) = CStr("'" & "" & ReadRs(adSummary, CurCol - 1))
           
         End Select
      Next
     
      If ArrayRow = RowPack Or CurRow = NoOfRecord + 1 + adSummary.RecordCount + 5 Then    '1 means the header row and 5 separate row between detail and summary
         ReDim tmpData(UBound(xlData, 2), UBound(xlData, 1))
         For i = 0 To UBound(xlData, 1)
            For J = 0 To UBound(xlData, 2)
               tmpData(J, i) = xlData(i, J)
            Next
         Next
         xlSheet1.Range(xlSheet1.Cells(StartRow, 1).Address, xlSheet1.Cells(CurRow, NoOfSum).Address).Value = tmpData
   
         ArrayRow = 0
         StartRow = CurRow + 1
         Erase tmpData
         ReDim xlData(NoOfSum, 0)
     End If
     UpdateStatus picStatus, CInt(wiStatus) + 80
     adSummary.MoveNext
     CurRow = CurRow + 1
   Loop
   
   adSummary.Close
   Set adSummary = Nothing
   UpdateStatus picStatus, 95
   
PrintExcel_Save:

   With xlSheet1.Range(xlSheet1.Cells(1, 1).Address, xlSheet1.Cells(CurRow - 1, NoOfFields).Address)
      .Borders(xlEdgeLeft).Weight = xlThin
      .Borders(xlEdgeTop).Weight = xlThin
      .Borders(xlEdgeBottom).Weight = xlThin
      .Borders(xlEdgeRight).Weight = xlThin
      .Borders(xlInsideVertical).Weight = xlThin
    End With
   
   If chkOnScreen Then
      With xlApp
         .Workbooks(1).Worksheets(1).Select
         .ShowToolTips = False
         .LargeButtons = False
         .ColorButtons = True
          DoEvents
         .Visible = True
         .WindowState = xlMaximized
         If Dir(lblDspSavePath) <> "" Then
             .Workbooks.Open lblDspSavePath
         End If
         UpdateStatus picStatus, 100, True
         Set xlSheet1 = Nothing
         Set xlApp = Nothing
      End With
   Else    'Output to file only / The default is no parameter
      xlApp.Workbooks(1).Worksheets(1).SaveAs lblDspSavePath.Caption
      xlApp.Quit
      Set xlSheet1 = Nothing
      Set xlApp = Nothing
      UpdateStatus picStatus, 100, True
      MsgBox lblDspSavePath.Caption & ""
   End If
   SaveUserDefault
   UpdateStatus picStatus, 0
   
   Exit Sub
   
Excel_Print_Err1:
   MsgBox "Excel_Print_Err1"
   On Error Resume Next
   UpdateStatus picStatus, 0
   Set xlSheet1 = Nothing
   Set xlApp = Nothing
   
   Exit Sub

Excel_Print_Err2:
   MsgBox "Excel_Print_Err2"
   On Error Resume Next
   UpdateStatus picStatus, 0
   xlApp.Workbooks("BOOK1").Saved = True
   xlApp.Workbooks.Close
   xlApp.Quit
   Set xlSheet1 = Nothing
   Set xlApp = Nothing

End Sub

Private Sub PrintListView()

   Dim wsSQL As String
   Dim wiCtr As Integer
   Dim wsFields As String
   Dim wsOrder As String
   Dim wsList() As String
   Dim NewfrmPrintList As New frmPrintList
   
   On Error GoTo PrintListView_Err
   Me.MousePointer = vbHourglass
   
   
   With lstSelect(PrintTo)
      ReDim wsList(.ListItems.Count, 2)
      For wiCtr = 1 To .ListItems.Count
         wsFields = wsFields & .ListItems(wiCtr).ListSubItems(ItemField).Text & ", "
         wsList(wiCtr, 1) = .ListItems(wiCtr).Text
         wsList(wiCtr, 2) = .ListItems(wiCtr).ListSubItems(ItemNumFlag).Text
      Next
   End With
   
   With lstSort(PrintTo)
      For wiCtr = 1 To .ListItems.Count
         wsOrder = wsOrder & .ListItems(wiCtr).ListSubItems(ItemField).Text & ", "
      Next
   End With
   wsFields = Mid$(wsFields, 1, Len(Trim(wsFields)) - 1)
   If Trim(wsOrder) <> "" Then
      wsOrder = Mid$(wsOrder, 1, Len(Trim(wsOrder)) - 1)
   End If

   If To_Value(txtTopN.Text) <> wiNoOfRecords Then
      wsSQL = " SELECT TOP " & To_Value(txtTopN.Text) & " WITH TIES "
      wsSQL = wsSQL & wsFields & " FROM RPT" & Me.TableID
   Else
      wsSQL = " SELECT " & wsFields & " FROM RPT" & Me.TableID
   End If
   wsSQL = wsSQL & " WHERE RPTUSRID = '" & Set_Quote(gsUserID) & "' "
   wsSQL = wsSQL & "   AND RPTDTETIM = '" & Change_SQLDate(Me.RptDteTim) & "' "
   If Trim(wsOrder) <> "" Then
      wsSQL = wsSQL & " ORDER BY " & wsOrder
   End If
      
   wsAction = cmdListView
   NewfrmPrintList.Fields = wsList
   NewfrmPrintList.NoOfCol = lstSelect(PrintTo).ListItems.Count
   NewfrmPrintList.RptTitle = Me.RptTitle
   NewfrmPrintList.Query = wsSQL
   NewfrmPrintList.Show vbModal
   
   Set NewfrmPrintList = Nothing
   SaveUserDefault
   Me.MousePointer = vbDefault
   
   Exit Sub
   
PrintListView_Err:
   Me.MousePointer = vbDefault
   MsgBox "PrintListView Err"
   

End Sub

Private Sub LoadListView()

   Dim wsSQL As String
   Dim adLayout As New ADODB.Recordset
   
   On Error GoTo LoadListView_Err
   Me.MousePointer = vbHourglass
   
   wsSQL = " SELECT COUNT(*) FROM sysLAYOUT "
   wsSQL = wsSQL & " WHERE LAYUSRID = '" & Set_Quote(gsUserID) & "' "
   wsSQL = wsSQL & " AND LAYPGMID = '" & Set_Quote(Me.TableID) & "' "
   adLayout.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
   If To_Value(adLayout(0).Value) > 0 Then
      adLayout.Close
      Set adLayout = Nothing
      LoadUserDefault
   Else
      adLayout.Close
      Set adLayout = Nothing
      LoadSPC
   End If
   
   Me.MousePointer = vbDefault
   Exit Sub
   
LoadListView_Err:
   'DISPLAY ERROR FUNCTION
   MsgBox "LoadListView Error!"
   Me.MousePointer = vbDefault
   On Error Resume Next
   adLayout.Close
   Set adLayout = Nothing

End Sub


Private Sub Ini_ListView()

   Dim wiCtr As Integer
   Dim clmX As columnheader
   'Dim wImg As ListImage
    
   ' Get ListView Small icons
   'With ImageList.ListImages
   '   Set wImg = .Add(1, "CHECKED", LoadResPicture("CHECKED", vbResIcon))
   '   Set wImg = .Add(2, "UNCHECKED", LoadResPicture("UNCHECKED", vbResIcon))
   'End With
   
   On Error GoTo Ini_ListView_Err
   
   'Initialize select list
   For wiCtr = PrintFrom To PrintTo
      With lstSelect(wiCtr)
         Set clmX = .ColumnHeaders. _
            Add(, , "Description", (.Width * 0.98))
         clmX.Alignment = lvwColumnLeft
         
         Set clmX = .ColumnHeaders. _
            Add(, , "Field", 0, lvwColumnCenter)
         clmX.Alignment = lvwColumnLeft
      
         Set clmX = .ColumnHeaders. _
            Add(, , "Type", 0, lvwColumnCenter)
         clmX.Alignment = lvwColumnLeft
      
         '.Icons = ImageList
         '.SmallIcons = ImageList
         .DragMode = 0
         .Sorted = False
         'SHOW THE CHECK BOX for the To-listview
         If wiCtr = PrintTo Then
            .CheckBoxes = True
         End If
      End With
   Next

   'Initialize sort list
   For wiCtr = PrintFrom To PrintTo
      With lstSort(wiCtr)
         Set clmX = .ColumnHeaders. _
            Add(, , "Description", (.Width * 0.98))
         clmX.Alignment = lvwColumnLeft
         
         Set clmX = .ColumnHeaders. _
            Add(, , "Field", 0, lvwColumnCenter)
         clmX.Alignment = lvwColumnLeft
             
         '.SmallIcons = ImageList
         '.Icons = ImageList
         .DragMode = 0
         .Sorted = False
         'SHOW THE CHECK BOX for the To-listview
         If wiCtr = PrintTo Then
            .CheckBoxes = True
         End If
      End With
   Next

   Set clmX = Nothing
   'Set wImg = Nothing
   
   Exit Sub
   
Ini_ListView_Err:
   MsgBox "Listview Error"

End Sub


Private Sub LoadSPC()

   Dim wsSQL As String
   Dim itmX As ListItem
   Dim itmY As ListItem
   Dim subX As ListSubItem
   Dim subY As ListSubItem
   Dim adSPC As New ADODB.Recordset
   
   On Error GoTo LoadSPC_Err
   
   wsSQL = " SELECT ScrFldID, ScrFldName, "
   wsSQL = wsSQL & " CASE WHEN USERTYPE IN (5, 6, 7, 8, 10, 11, 21, 24) THEN 'N' "
   wsSQL = wsSQL & " WHEN USERTYPE IN (12, 22, 80) THEN 'D' "
   wsSQL = wsSQL & " WHEN USERTYPE IN (19) THEN 'T' "
   wsSQL = wsSQL & " ELSE 'C' END AS ScrFldType FROM sysScrCaption, SYSCOLUMNS "
   wsSQL = wsSQL & " WHERE ScrType = 'FIL' "
   wsSQL = wsSQL & " AND SYSCOLUMNS.ID = OBJECT_ID('RPT" & Me.TableID & "') "
   wsSQL = wsSQL & " AND SYSCOLUMNS.NAME = ScrFldID "
   wsSQL = wsSQL & " AND ScrPgmID = '" & Set_Quote(Me.TableID) & "' "
   wsSQL = wsSQL & " AND ScrLangID = '" & gsLangID & "' "
   wsSQL = wsSQL & " AND ISNULL(RTRIM(ScrFldID), '') <> '' "
   wsSQL = wsSQL & " ORDER BY ScrSeqNo "
   adSPC.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
   If adSPC.RecordCount = 0 Then
      wsSQL = " SELECT SYSCOLUMNS.NAME AS ScrFldID, SYSCOLUMNS.NAME AS ScrFldName, "
      wsSQL = wsSQL & " CASE WHEN USERTYPE IN (5, 6, 7, 8, 10, 11, 21, 24) THEN 'N' "
      wsSQL = wsSQL & " WHEN USERTYPE IN (12, 22, 80) THEN 'D' "
      wsSQL = wsSQL & " WHEN USERTYPE IN (19) THEN 'T' "
      wsSQL = wsSQL & " ELSE 'C' END AS ScrFldType FROM SYSCOLUMNS "
      wsSQL = wsSQL & " WHERE SYSCOLUMNS.ID = OBJECT_ID('RPT" & Me.TableID & "') "
      wsSQL = wsSQL & " ORDER BY SYSCOLUMNS.NAME "
      adSPC.Close
      adSPC.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
      If adSPC.RecordCount = 0 Then
         MsgBox "No " & "RPT" & Me.TableID & "in System"
         GoTo LoadSPC_Exit
      End If
   End If
   
   Do Until adSPC.EOF
      
      If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
         Set itmX = lstSelect(PrintTo).ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
         Set itmY = lstSort(PrintFrom).ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
      Else
         Set itmX = lstSelect(PrintTo).ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
         Set itmY = lstSort(PrintFrom).ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
      End If
      
      With itmX
         Set subX = itmX.ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
         Set subY = itmX.ListSubItems.Add(ItemNumFlag, , Trim(ReadRs(adSPC, "ScrFldType")))
         '.SubItems(ItemField) = adSPC("SPCFLDID").Value
         '.SubItems(ItemNumFlag) = adSPC("SPCMNUSPA").Value
         '.SmallIcon = 2
         '.Icon = 2
      End With
      Set subY = Nothing
      
      Set subY = itmY.ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFLDID")))
      'itmY.SubItems(ItemField) = adSPC("SPCFLDID").Value

      adSPC.MoveNext
   Loop
   
   adSPC.Close
   Set adSPC = Nothing
   Set itmX = Nothing
   Set itmY = Nothing
   Set subX = Nothing
   Set subY = Nothing
   
   Exit Sub
   
LoadSPC_Err:
   'DISPLAY ERROR FUNCTION
   MsgBox "LoadSPC Err!"
   
LoadSPC_Exit:
   On Error Resume Next
   adSPC.Close
   Set adSPC = Nothing
   Set itmX = Nothing

End Sub

Private Sub LoadUserDefault()

   'Load the user default (previous) fields selection for list and index

   Dim wsSQL As String
   Dim wsSql2 As String
   Dim wsSql3 As String
   Dim itmX As ListItem
   Dim subX As ListSubItem
   Dim subY As ListSubItem
   Dim adSPC As New ADODB.Recordset
   
   On Error GoTo LoadDefault_Err
   
   
   wsSQL = " SELECT ScrFldID, ScrFldName, "
   wsSQL = wsSQL & "CASE WHEN USERTYPE IN (5, 6, 7, 8, 10, 11, 21, 24) THEN 'N' "
   wsSQL = wsSQL & " WHEN USERTYPE IN (12, 22, 80) THEN 'D' "
   wsSQL = wsSQL & " WHEN USERTYPE IN (19) THEN 'T' "
   wsSQL = wsSQL & " ELSE 'C' END AS ScrFldType, "
   wsSql3 = wsSQL
   wsSQL = wsSQL & " LayTotal, LayGroup, LaySelect, LaySort, LayOnScreen "
   wsSQL = wsSQL & " FROM sysScrCaption, sysLayout, SYSCOLUMNS "
   wsSQL = wsSQL & " WHERE ScrType = 'FIL' "
   wsSQL = wsSQL & " AND SYSCOLUMNS.ID = OBJECT_ID('RPT" & Me.TableID & "') "
   wsSQL = wsSQL & " AND SYSCOLUMNS.NAME = ScrFLDID "
   wsSQL = wsSQL & " AND ScrPgmID = '" & Set_Quote(Me.TableID) & "' "
   wsSQL = wsSQL & " AND LayUsrID = '" & Set_Quote(gsUserID) & "' "
   wsSQL = wsSQL & " AND ScrPgmID = LayPgmID "
   wsSQL = wsSQL & " AND LayFldID = ScrFLDID "
   wsSQL = wsSQL & " AND ScrLangID = '" & gsLangID & "' "
   wsSQL = wsSQL & " AND ISNULL(RTRIM(ScrFLDID), '') <> '' "
   wsSql2 = wsSQL & " AND LaySelect <> 999 "
   wsSql2 = wsSql2 & " ORDER BY LaySort, LaySelect "   'USED FOR LSTSORT
   wsSQL = wsSQL & " ORDER BY LaySelect, ScrSeqNo "     'USED FOR LSTSELECT
   adSPC.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
   If adSPC.RecordCount = 0 Then
      LoadSPC
      GoTo LoadDefault_Exit
   End If
   
   chkOnScreen.Value = ReadRs(adSPC, "LayOnScreen")
'   wiIndex = 0
   Do Until adSPC.EOF
      'IF THE VALUE IS 999 THEN PUT INTO LISTVIEW FROM
      If To_Value(ReadRs(adSPC, "LaySelect")) = 999 Then
         With lstSelect(PrintFrom)
            If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
            Else
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
            End If
            
            With itmX
               Set subX = .ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
               Set subY = .ListSubItems.Add(ItemNumFlag, , Trim(ReadRs(adSPC, "ScrFldType")))
               '.SubItems(ItemField) = adSPC("SPCFLDID").Value
               '.SubItems(ItemNumFlag) = adSPC("SPCMNUSPA").Value
            End With
   
         End With
      Else  'ELSE PUT INTO LISTVIEW TO
         With lstSelect(PrintTo)
            If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
            Else
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
            End If
         
            With itmX
               If UCase(ReadRs(adSPC, "LayTotal")) = "Y" Then
                  .Checked = True
               Else
                  .Checked = False
               End If
               Set subX = .ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
               Set subY = .ListSubItems.Add(ItemNumFlag, , Trim(ReadRs(adSPC, "ScrFldType")))
               '.SubItems(ItemField) = adSPC("SPCFLDID").Value
               '.SubItems(ItemNumFlag) = adSPC("SPCMNUSPA").Value
            End With
         End With
      End If
      adSPC.MoveNext
   Loop
   adSPC.Close
   Set adSPC = Nothing
   Set itmX = Nothing
   Set subX = Nothing
   Set subY = Nothing
   
   adSPC.Open wsSql2, cnCon, adOpenStatic, adLockOptimistic
   
   If adSPC.RecordCount = 0 Then
      GoTo LoadDefault_Exit
   End If
   
'   wiIndex = 0
   Do Until adSPC.EOF
      'IF THE TO POSITION VALUE IS 999 THEN PUT INTO LISTVIEW FROM
      If To_Value(adSPC("LaySort")) = 999 Then
         With lstSort(PrintFrom)
            If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
            Else
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
            End If
            
            With itmX
               Set subX = .ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
               '.SubItems(ItemField) = adSPC("SPCFLDID").Value
            End With
   
         End With
      Else  'ELSE PUT INTO LISTVIEW TO
         With lstSort(PrintTo)
            If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
            Else
               Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
            End If
         
            With itmX
               If UCase(ReadRs(adSPC, "LayGroup")) = "Y" Then
                  .Checked = True
               Else
                  .Checked = False
               End If
               Set subX = .ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
               '.SubItems(ItemField) = adSPC("SPCFLDID").Value
            End With
         End With
      End If
      adSPC.MoveNext
   Loop
   adSPC.Close
   Set adSPC = Nothing
   Set itmX = Nothing
   Set subY = Nothing
   
   wsSql3 = wsSql3 & " ScrSeqNo FROM sysScrCaption, SYSCOLUMNS "
   wsSql3 = wsSql3 & " WHERE ScrType = 'FIL' "
   wsSql3 = wsSql3 & " AND SYSCOLUMNS.ID = OBJECT_ID('RPT" & Me.TableID & "') "
   wsSql3 = wsSql3 & " AND SYSCOLUMNS.NAME = ScrFLDID "
   wsSql3 = wsSql3 & " AND ScrPgmID = '" & Set_Quote(Me.TableID) & "' "
   wsSql3 = wsSql3 & " AND NOT EXISTS ( SELECT NULL FROM sysLAYOUT"
   wsSql3 = wsSql3 & " WHERE ScrPGMID = LAYPGMID "
   wsSql3 = wsSql3 & " AND ScrFLDID = LAYFLDID "
   wsSql3 = wsSql3 & " AND LAYUSRid = '" & Set_Quote(gsUserID) & "') "
   wsSql3 = wsSql3 & " ORDER BY Scrseqno "
   
   adSPC.Open wsSql3, cnCon, adOpenStatic, adLockOptimistic
   
   If adSPC.RecordCount = 0 Then GoTo LoadDefault_Exit
   
   Do Until adSPC.EOF
      With lstSelect(PrintFrom)
         If Trim$(ReadRs(adSPC, "ScrFldName")) = "" Then
            Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldID")))
         Else
            Set itmX = .ListItems.Add(, , Trim(ReadRs(adSPC, "ScrFldName")))
         End If
         
         With itmX
            Set subX = .ListSubItems.Add(ItemField, , Trim(ReadRs(adSPC, "ScrFldID")))
            Set subY = .ListSubItems.Add(ItemNumFlag, , Trim(ReadRs(adSPC, "ScrFldType")))
            '.SubItems(ItemField) = adSPC("SPCFLDID").Value
         End With
      End With
      adSPC.MoveNext
   Loop
   adSPC.Close
   Set adSPC = Nothing
   
   Set itmX = Nothing
   Set subX = Nothing
   
   Exit Sub
   
LoadDefault_Err:
   MsgBox "LoadDeafult Err!"
   
LoadDefault_Exit:
   On Error Resume Next
   adSPC.Close
   Set adSPC = Nothing

End Sub


Private Sub MoveSelectItem(ByVal Index As Integer, ByVal wiAll As Boolean)

   Dim wiCtr As Integer
   Dim wiFrom As Integer
   Dim wiTo As Integer
   Dim itmX As ListItem
   Dim itmY As ListItem
   Dim subX As ListSubItem
   Dim subY As ListSubItem
   Dim wiCtr2 As Integer
   Dim wiSelected As Integer
   Dim wiDeleted As Integer
   Dim wiIndex As Integer
   Dim wsField As String
   
   On Error GoTo MoveSelectItem_Err
   
   If Index = PrintFrom Then
      wiFrom = PrintFrom
      wiTo = PrintTo
   Else
      wiTo = PrintFrom
      wiFrom = PrintTo
   End If
   
   wiSelected = False
   With lstSelect(wiFrom)
      For wiCtr = 1 To .ListItems.Count
      
         'Add the selected or all items to the another list view
         If .ListItems(wiCtr).Selected Or wiAll = True Then
            
            wiSelected = True
            wsField = .ListItems(wiCtr).Text
            Set itmX = lstSelect(wiTo).ListItems.Add(, , .ListItems(wiCtr).Text)
                    
            With itmX
               Set subX = .ListSubItems.Add(ItemField, , lstSelect(wiFrom).ListItems(wiCtr).ListSubItems(ItemField).Text)
               Set subY = .ListSubItems.Add(ItemNumFlag, , lstSelect(wiFrom).ListItems(wiCtr).ListSubItems(ItemNumFlag).Text)
               '.SubItems(ItemField) = lstSelect(wiFrom).ListItems(wiCtr).SubItems(ItemField)
               '.SubItems(ItemNumFlag) = lstSelect(wiFrom).ListItems(wiCtr).SubItems(ItemNumFlag)
            End With
            Set subX = Nothing
            Set subY = Nothing
            
            'if the adding listview is PrintTo, add the item to the lstSort also.
            If wiTo = PrintTo Then
               Set itmY = lstSort(PrintFrom).ListItems.Add(, , .ListItems(wiCtr).Text)
            
               With itmY
                  Set subY = .ListSubItems.Add(ItemField, , lstSelect(wiFrom).ListItems(wiCtr).ListSubItems(ItemField).Text)
                  '.SubItems(ItemField) = lstSelect(wiFrom).ListItems(wiCtr).SubItems(ItemField)
               End With
               Set subY = Nothing
            'else remove the item
            Else
               ' locate the lstsort index
               If wiAll = True Then
                  lstSort(PrintFrom).ListItems.Clear
                  lstSort(PrintTo).ListItems.Clear
               Else
                  With lstSort(PrintFrom)
                     wiDeleted = False
                     For wiCtr2 = 1 To .ListItems.Count
                        If .ListItems(wiCtr2).Text = wsField Then
                           wiDeleted = True
                           .ListItems.Remove wiCtr2
                           Exit For
                        End If
                     Next
                  End With
                  'if not found at from listview, find that item at to list view
                  If wiDeleted = False Then
                     With lstSort(PrintTo)
                        For wiCtr2 = 1 To .ListItems.Count
                           If .ListItems(wiCtr2).Text = wsField Then
                              .ListItems.Remove wiCtr2
                              Exit For
                           End If
                        Next
                     End With
                  End If
               End If
            End If
            
         End If
      Next
      'Remove selected or all items from the current list view
      If wiAll = True Then
         .ListItems.Clear
      Else
         If wiSelected = False Then Exit Sub
         .ListItems.Remove .SelectedItem.Index
      End If
            
   End With
   
   Set itmX = Nothing
   Set itmY = Nothing
   
   Exit Sub
   
MoveSelectItem_Err:
   MsgBox "Error in MoveSelectItem!"
   
End Sub

Private Sub MoveSortItem(ByVal Index As Integer, ByVal wiAll As Boolean)

   Dim wiCtr As Integer
   Dim wiFrom As Integer
   Dim wiTo As Integer
   Dim itmX As ListItem
   Dim itmY As ListItem
   Dim subX As ListSubItem
   Dim wiCtr2 As Integer
   Dim wiSelected As Integer
   Dim wiDeleted As Integer
   Dim wiIndex As Integer
   Dim wsField As String
   
   On Error GoTo MoveSortItem_Err
   
   If Index = PrintFrom Then
      wiFrom = PrintFrom
      wiTo = PrintTo
   Else
      wiTo = PrintFrom
      wiFrom = PrintTo
   End If
   
   wiSelected = False
   With lstSort(wiFrom)
      For wiCtr = 1 To .ListItems.Count
      
         'Add the selected or all items to the another list view
         If .ListItems(wiCtr).Selected Or wiAll = True Then
            
            wiSelected = True
            wsField = .ListItems(wiCtr).Text
            Set itmX = lstSort(wiTo).ListItems.Add(, , .ListItems(wiCtr).Text)
                    
            With itmX
               Set subX = .ListSubItems.Add(ItemField, , lstSort(wiFrom).ListItems(wiCtr).ListSubItems(ItemField).Text)
               '.SubItems(ItemField) = lstSort(wiFrom).ListItems(wiCtr).SubItems(ItemField)
            End With
                       
         End If
      Next
      'Remove selected or all items from the current list view
      If wiAll = True Then
         .ListItems.Clear
      Else
         If wiSelected = False Then Exit Sub
         .ListItems.Remove .SelectedItem.Index
      End If
            
   End With
   
   Set itmX = Nothing
   Set itmY = Nothing
   Set subX = Nothing
   
   Exit Sub

MoveSortItem_Err:
  MsgBox "Error in MOveSortItem!"
   
End Sub

Private Sub MovePosition(ByVal inDirection As Integer, _
                         ByVal inListView As Integer)

   Dim woList As ListView
   Dim itmX As ListItem
   Dim subX As ListSubItem
   Dim subY As ListSubItem
   Dim wiCtr As Integer
   Dim Index As Integer
   Dim wsDesc As String
   Dim wsField As String
   Dim wsNumFlag As String
   Dim wiChecked As Integer
   
   On Error GoTo MovePosition_Err
   
   If inListView = TabSelect Then
      Set woList = lstSelect(PrintTo)
   Else
      Set woList = lstSort(PrintTo)
   End If
   
   With woList
      'locate the 1st item's position, if no selected item found then exit
      Index = .SelectedItem.Index
      If Index < 1 Or Index > .ListItems.Count Then GoTo MovePosition_Exit
      
      'if the index = top of the list and move up and vice versa => exit
      If (Index = 1 And inDirection = MoveUp) Or _
         (Index = .ListItems.Count And inDirection = MoveDown) Then
         GoTo MovePosition_Exit
      End If
      
      'Move the position
      wsDesc = .ListItems(Index).Text
      wsField = .ListItems(Index).ListSubItems(ItemField).Text
      If inListView = TabSelect Then
         wsNumFlag = .ListItems(Index).ListSubItems(ItemNumFlag).Text
      End If
      wiChecked = .ListItems(Index).Checked
      .ListItems.Remove Index
      If inDirection = MoveUp Then
         Index = Index - 1
      Else
         Index = Index + 1
      End If
      Set itmX = .ListItems.Add(Index, , wsDesc)
      
      With itmX
         .Checked = wiChecked
         Set subX = .ListSubItems.Add(ItemField, , wsField)
         If inListView = TabSelect Then
            Set subY = .ListSubItems.Add(ItemNumFlag, , wsNumFlag)
         End If
         '.Icon = 1
         '.SmallIcon = 1
      End With
      
      .ListItems(Index).Selected = True
      .SetFocus
      
   End With
   
   Set itmX = Nothing
   Set subX = Nothing
   Set subY = Nothing
   Set woList = Nothing
   Exit Sub
   
MovePosition_Err:
  MsgBox "Error in MovePosition!"
MovePosition_Exit:
   Set itmX = Nothing
   Set subX = Nothing
   Set subY = Nothing
   Set woList = Nothing
   Exit Sub

End Sub

Private Sub SaveUserDefault()

   Dim wsSQL As String
   Dim Index As Integer
   Dim wiCtr As Integer
   Dim wiCtr2 As Integer
   Dim wiIdxPst As Integer
   Dim wsLupDte As String
   Dim cmdSPC As New ADODB.Command
   Dim adSPC As New ADODB.Recordset
   
   On Error GoTo SaveUserDefault_Err
   
   Set cmdSPC.ActiveConnection = cnCon
   cmdSPC.CommandText = "USP_SysLayout"
   cmdSPC.CommandType = adCmdStoredProc
   cmdSPC.Parameters.Refresh
   
 '  wsLupDte = Change_SQLDate(Me.RptDteTim)
   wsLupDte = Change_SQLDate(Now)
   
   Index = 1
   With lstSelect(PrintFrom)
      For wiCtr = 1 To .ListItems.Count
         wsLupDte = wsLupDte
         SetSPPara cmdSPC, 1, Index
         SetSPPara cmdSPC, 2, gsUserID
         SetSPPara cmdSPC, 3, wsLupDte
         SetSPPara cmdSPC, 4, Me.TableID
         SetSPPara cmdSPC, 5, .ListItems(wiCtr).ListSubItems(ItemField).Text
         SetSPPara cmdSPC, 6, 999
         SetSPPara cmdSPC, 7, 999
         SetSPPara cmdSPC, 8, "N"
         SetSPPara cmdSPC, 9, "N"
         SetSPPara cmdSPC, 10, chkOnScreen.Value
         'SetSPPara cmdSPC, 11, "RPT" & Me.TableID
         cmdSPC.Execute
         Index = Index + 1
      Next
   End With
   
   With lstSelect(PrintTo)
      For wiCtr = 1 To .ListItems.Count
         SetSPPara cmdSPC, 1, Index
         SetSPPara cmdSPC, 2, gsUserID
         SetSPPara cmdSPC, 3, wsLupDte
         SetSPPara cmdSPC, 4, Me.TableID
         SetSPPara cmdSPC, 5, .ListItems(wiCtr).ListSubItems(ItemField).Text
         SetSPPara cmdSPC, 6, wiCtr
         'find the field position in lstSort(to)
         wiIdxPst = 999
         For wiCtr2 = 1 To lstSort(PrintTo).ListItems.Count
            If lstSort(PrintTo).ListItems(wiCtr2).ListSubItems(ItemField).Text = .ListItems(wiCtr).ListSubItems(ItemField).Text Then
               wiIdxPst = wiCtr2
               Exit For
            End If
         Next
         SetSPPara cmdSPC, 7, wiIdxPst
         SetSPPara cmdSPC, 8, IIf(.ListItems(wiCtr).Checked, "Y", "N")
         If wiIdxPst = 999 Then
            SetSPPara cmdSPC, 9, "N"
         Else
            SetSPPara cmdSPC, 9, IIf(lstSort(PrintTo).ListItems(wiIdxPst).Checked, "Y", "N")
         End If
         SetSPPara cmdSPC, 10, chkOnScreen.Value
         'SetSPPara cmdSPC, 11, "RPT" & Me.TableID
         cmdSPC.Execute
         Index = Index + 1
      Next
   End With
   
   Exit Sub

SaveUserDefault_Err:
   MsgBox "Error in SaveUserDefault!"
   MsgBox Err.Description
End Sub

Private Function Init_Excel(woExcelSheet1 As Object, _
                    xlApp2 As Object) As Integer

   Dim xlSheet2 As Object

' Variables used to format the selection criteria
' within the MS Excel Header and Footer
   Dim wsLeftHdr As String
   Dim wsCenterHdr As String
   Dim wsRightHdr As String
   Dim wsLeftFtr As String
   Dim wsRightFtr As String
   Dim wsTxt As String
   Dim wsLeftTxt As String
   Dim wsMid As String
   Dim wiLenHdr As Integer
   Dim wiLenFtr As Integer
   Dim wiFor As Integer
   Dim wiCtr As Integer
   Dim wiOldCtr As Integer
   Dim wbOk As Boolean
   Dim wbDot As Boolean
   Dim wbTxtEmpty As Boolean
' End of variables

   Init_Excel = False

   'GET THE REQUIRED DATA FROM RPTHDR
   On Error GoTo Excel_Err
   woExcelSheet1.PageSetup.PrintTitleRows = woExcelSheet1.Rows(1).Address

   ' ************************** WORKSHEET (DATA) ************************** '
   'PAGE SETUP
   With woExcelSheet1.PageSetup
      'CREATE FOOTER
      wsLeftFtr = "&""Times New Roman""&6"
      wsRightFtr = "&""Times New Roman""&8Page &P of &N"
      wiLenFtr = Len(wsLeftFtr) + Len(wsRightFtr)
      
      'CREATE HEADER
      wsRightHdr = "&""Times New Roman""&8USER: " & gsUserID & vbLf & "   DATE: " & Change_SQLDate(Me.RptDteTim)
      wsCenterHdr = "&""Times New Roman,Bold""&10" & Me.RptTitle
      wsLeftHdr = "&""Times New Roman""&8" & gsComNam & "&6" & vbLf & vbLf
      wiLenHdr = Len(wsRightHdr) + Len(wsCenterHdr) + Len(wsLeftHdr)
      
      'wbDot = False
      'wbTxtEmpty = True
      'wiCtr = 0
      'wiOldCtr = 0
      'wsLeftTxt = ""
      'wsTxt = Trim(lblSel01.Tag)
      'If wsTxt <> "" Then
      '   For wiFor = 1 To Len(wsTxt)
      '      wsMid = Mid(wsTxt, wiFor, 1)
      '      wbOk = False
      '
      '      ' Check linefeed
      '      If wsMid <> vbLf Then
      '         wsLeftTxt = wsLeftTxt & wsMid
      '         wiCtr = wiCtr + 1
      '      Else
      '         If Trim(wsLeftTxt) <> "" Then
      '            wbTxtEmpty = False
      '            ' Concatenate to Left Header
      '            If wiLenHdr + wiCtr <= 240 Then
      '               wsLeftHdr = wsLeftHdr & wsLeftTxt & Space(5)
      '               wiOldCtr = wiCtr + 5
      '               wbOk = True
      '            Else
      '               ' Concatenate to Left Footer
      '               If wiLenFtr + (wiCtr - wiOldCtr) <= 230 Then
      '                  wsLeftFtr = wsLeftFtr & wsLeftTxt & Space(5)
      '                  wbOk = True
      '               Else
      '                  wbDot = True
      '                  Exit For
      '               End If
      '            End If
      '
      '            ' New Selection Criteria
      '            If wbOk Then
      '               wiCtr = wiCtr + 5
      '               wsLeftTxt = ""
      '            End If
      '         Else
      '            wiCtr = wiCtr - Len(wsLeftTxt)
      '         End If
      '      End If
      '   Next wiFor
      'End If
        
      'SET HEADER
      .LeftHeader = wsLeftHdr
      .CenterHeader = wsCenterHdr
      .RightHeader = wsRightHdr
      
      'SET FOOTER
      .LeftFooter = wsLeftFtr & IIf(wbDot = True, String(10, "."), "")
      .CenterFooter = ""
      .RightFooter = wsRightFtr
      
      'SET ALIGNMENT
      .CenterHorizontally = True
      .CenterVertically = False
      .TopMargin = xlApp2.InchesToPoints(0.78)
      .BottomMargin = xlApp2.InchesToPoints(0.35)
      .LeftMargin = xlApp2.InchesToPoints(0.2)
      .RightMargin = xlApp2.InchesToPoints(0.2)
      .HeaderMargin = xlApp2.InchesToPoints(0.2)
      .FooterMargin = xlApp2.InchesToPoints(0.2)
      
      'SET PAGE ORIENTATION
      .Orientation = xlLandscape

   End With

   ''GENERAL CELL/S SETUP
   'With woExcelSheet1
   '   .Cells.Select
   '   .Cells.Font.FontStyle = "Regular"
   '   .Cells.Font.Name = "Times New Roman"
   '   .Cells.Font.Size = 8
   '   .Cells.Borders.Weight = xlThin
   '   .Range("A1").Select
   'End With

' ******************* WORKSHEET (SELECTION CRITERIA) ******************* '

   'If Not wbTxtEmpty Then
   If UBound(Me.Selection) > 0 Then
      With xlApp2.Workbooks(1)
         .Worksheets.Add.Move AFTER:=.Worksheets(.Worksheets.Count)
         Set xlSheet2 = .Worksheets(.Worksheets.Count)
      End With
    
      'PAGE SETUP
      With xlSheet2.PageSetup
         'SET HEADER
         .LeftHeader = "&""Times New Roman""&8" & gsComNam
         .CenterHeader = woExcelSheet1.PageSetup.CenterHeader & vbLf & vbLf
         .CenterHeader = .CenterHeader & "&""Times New Roman,Bold""&9" & "SELECTION CRITERIA"
         .RightHeader = woExcelSheet1.PageSetup.RightHeader
         
         'SET FOOTER
         .LeftFooter = ""
         .CenterFooter = ""
         .RightFooter = "&""Times New Roman""&8Page &P of &N"
         
         'SET ALIGNMENT
         .CenterHorizontally = woExcelSheet1.PageSetup.CenterHorizontally
         .CenterVertically = woExcelSheet1.PageSetup.CenterVertically
         .TopMargin = woExcelSheet1.PageSetup.TopMargin
         .BottomMargin = woExcelSheet1.PageSetup.BottomMargin
         .LeftMargin = woExcelSheet1.PageSetup.LeftMargin
         .RightMargin = woExcelSheet1.PageSetup.RightMargin
         .HeaderMargin = woExcelSheet1.PageSetup.HeaderMargin
         .FooterMargin = woExcelSheet1.PageSetup.FooterMargin
         
         'SET PAGE ORIENTATION
         .Orientation = woExcelSheet1.PageSetup.Orientation
      End With

      'wiCtr = 0
      'wsLeftTxt = ""
'      wsTxt = Trim(lblSel01.Tag)
      For wiCtr = 1 To UBound(Me.Selection)
'         wsMid = Mid(wsTxt, wiFor, 1)
        
         ' Check linefeed
         'If wsMid <> vbLf Then
         '   wsLeftTxt = wsLeftTxt & wsMid
         'Else
         '   If Trim(wsLeftTxt) <> "" Then
         '      wiCtr = wiCtr + 1
         With xlSheet2
            .Cells(wiCtr, 1).Value = Me.Selection(wiCtr)
            .Cells(wiCtr, 1).EntireColumn.AutoFit
         End With
         '      wsLeftTxt = ""
         '   End If
         'End If
      Next
      
      If To_Value(txtTopN.Text) <> wiNoOfRecords Then
         xlSheet2.Cells(wiCtr + 1, 1).Value = "'" & lblTopN.Caption & " " & To_Value(txtTopN.Text)
      End If

      'GENERAL CELL/S SETUP
      With xlSheet2
         .Cells.Font.FontStyle = "Regular"
         .Cells.Font.Name = "Times New Roman"
         .Cells.Font.Size = 8
         .Cells.Borders.Weight = xlThin
         .Protect
      End With
   End If

   Set xlSheet2 = Nothing
   Init_Excel = True

   Exit Function

Excel_Err:
   MsgBox "Exel_err"
   On Error Resume Next
   Set xlSheet2 = Nothing
   
End Function

Private Sub Ini_Caption()

    Call Get_Scr_Item("FRMPRINT", waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
   'Me.Caption = Get_Caption(waScrItm, "SCRHDR") & " - " & Me.RptTitle
   
   'fmePrnOpt.Caption = Get_Caption(waScrItm, "PRINTOPT")
   'optPrint(cmdPreview).Caption = Get_Caption(waScrItm, "CMDPREVIEW")
   'optPrint(cmdPrinter).Caption = Get_Caption(waScrItm, "CMDPRINTER")
   'optPrint(cmdExcel).Caption = Get_Caption(waScrItm, "CMDEXCEL")
   'optPrint(cmdListView).Caption = Get_Caption(waScrItm, "CMDLISTVIEW")
   'cmdPrint(PrintOK).Caption = Get_Caption(waScrItm, "PRINTOK")
   'cmdPrint(PrintCancel).Caption = Get_Caption(waScrItm, "PRINTCANCEL")
   'cmdPrint(PrintSave).Caption = Get_Caption(waScrItm, "PRINTSAVE")
   'cmdPrint(PrintPrinter).Caption = Get_Caption(waScrItm, "PRINTPRINTER")
   'cmdPrint(PrintDetail).Caption = Get_Caption(waScrItm, "PRINTDETAIL")
   'lblSavePath.Caption = Get_Caption(waScrItm, "LBLSAVEPATH")
   'lblNoOfRecords.Caption = Get_Caption(waScrItm, "LBLNOOFRECORDS")
   'tabFieldSelect.TabCaption(TabSelect) = Get_Caption(waScrItm, "TABSELECT")
   'tabFieldSelect.TabCaption(TabSort) = Get_Caption(waScrItm, "TABSORT")
   'cmdSort(MoveRight).Caption = Get_Caption(waScrItm, "MOVERIGHT")
   'cmdSort(MoveRightAll).Caption = Get_Caption(waScrItm, "MOVERIGHTALL")
   'cmdSort(MoveLeft).Caption = Get_Caption(waScrItm, "MOVELEFT")
   'cmdSort(MoveLeftAll).Caption = Get_Caption(waScrItm, "MOVELEFTALL")
   'lblTopN.Caption = Get_Caption(waScrItm, "TOPN")
   'chkOnScreen.Caption = Get_Caption(waScrItm, "ONSCREEN")
   
    Me.Caption = Get_Caption(waScrItm, "SCRHDR") & " - " & Me.RptTitle
    lblSavePath.Caption = Get_Caption(waScrItm, "SAVEPATH")
    lblNoOfRecords.Caption = Get_Caption(waScrItm, "NOOFRECORDS")
    tabFieldSelect.TabCaption(TabSelect) = Get_Caption(waScrItm, "TABSELECT")
    tabFieldSelect.TabCaption(TabSort) = Get_Caption(waScrItm, "TABSORT")
    cmdSort(MoveRight).Caption = Get_Caption(waScrItm, "MOVERIGHT")
    cmdSort(MoveRightAll).Caption = Get_Caption(waScrItm, "MOVERIGHTALL")
    cmdSort(MoveLeft).Caption = Get_Caption(waScrItm, "MOVELEFT")
    cmdSort(MoveLeftAll).Caption = Get_Caption(waScrItm, "MOVELEFTALL")
    cmdSort(MoveUp).Caption = Get_Caption(waScrItm, "MOVEUP")
    cmdSort(MoveDown).Caption = Get_Caption(waScrItm, "MOVEDOWN")
    cmdSelect(MoveUp).Caption = Get_Caption(waScrItm, "SELECTMOVEUP")
    cmdSelect(MoveDown).Caption = Get_Caption(waScrItm, "SELECTMOVEDOWN")
    lblTopN.Caption = Get_Caption(waScrItm, "TOPN")
    chkOnScreen.Caption = Get_Caption(waScrItm, "ONSCREEN")
    
    tbrProcess.Buttons(tcPreview).ToolTipText = Get_Caption(waScrToolTip, tcPreview) & "(F2)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F3)"
    tbrProcess.Buttons(tcExcel).ToolTipText = Get_Caption(waScrToolTip, tcExcel) & "(F5)"
    tbrProcess.Buttons(tcBrowse).ToolTipText = Get_Caption(waScrToolTip, tcBrowse) & "(F6)"
    tbrProcess.Buttons(tcPrinter).ToolTipText = Get_Caption(waScrToolTip, tcPrinter) & "(F9)"
    tbrProcess.Buttons(tcDetail).ToolTipText = Get_Caption(waScrToolTip, tcDetail) & "(F10)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
End Sub



Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "None"
            With tbrProcess
                .Buttons(tcPreview).Enabled = False
                .Buttons(tcPrint).Enabled = False
                .Buttons(tcExcel).Enabled = False
                .Buttons(tcBrowse).Enabled = False
                .Buttons(tcPrinter).Enabled = False
               End With
            
          Case "All"
            With tbrProcess
                .Buttons(tcPreview).Enabled = True
                .Buttons(tcPrint).Enabled = True
                .Buttons(tcExcel).Enabled = True
                .Buttons(tcBrowse).Enabled = True
                .Buttons(tcPrinter).Enabled = True
               End With
            
        
        Case "NoPrint"
              With tbrProcess
                .Buttons(tcPreview).Enabled = False
                .Buttons(tcPrint).Enabled = False
                .Buttons(tcPrinter).Enabled = False
               End With
           
         Case "Print"
              With tbrProcess
                .Buttons(tcPreview).Enabled = True
                .Buttons(tcPrint).Enabled = True
                .Buttons(tcPrinter).Enabled = True
                End With
           
    End Select
End Sub

Private Sub FormExit()
If tbrProcess.Buttons(tcPrint).Enabled = True Then
         Unload Me
         Exit Sub
         'Me.Hide
      Else
         If wsAction = cmdFail Then
            Unload Me
            Exit Sub
         Else
            wdbCon.Cancel
            UpdateStatus picStatus, 0
            Unload Me
            Exit Sub
         End If
         'Me.Hide
      End If
     
End Sub


Private Sub PrinterSetup()

    On Error GoTo PrintErr
      
      cdPrinter.ShowPrinter
      
PrintErr:
   Me.MousePointer = vbDefault
   If cdPrinter.CancelError = True Then
      Exit Sub
   End If
      
End Sub


Private Sub PrintDetailForm()

 If Me.WindowState = 2 Then Exit Sub ' if maximize no need to change the screen size
      If Me.Height = SummaryHeight Then
         Me.Height = DetailHeight
      Else
         Me.Height = SummaryHeight
End If
End Sub

Private Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase)
'-----------------------------------------------------------
' SUB: UpdStatusBar
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
    Dim intPercent As Double
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    If IsMissing(fBorderCase) Then fBorderCase = False
    
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &HFF0000 ' blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    intPercent = sngPercent
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    If sngPercent <> 0 Then
       strPercent = Format$(intPercent) & "%"
    Else
       strPercent = ""
    End If
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
       pic.Line (0, 0)-(pic.Width * (sngPercent / 100), pic.Height), pic.ForeColor, BF
    Else
       pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If
    
    pic.Refresh

End Sub


Private Sub Access_Print(inMode As Integer)
'   Declare Database variable and Microsoft Access Form variable.
    Dim wsPara As String
'   Dim appAccess As Access.Application
    Dim appAccess As Object
    Dim tmpString As String
    Dim accessHandle As Long

    On Error GoTo Access_Print_Err
    
    UpdateStatus picStatus, 10
    
'   Return instance of Application object.
    MousePointer = vbHourglass
    
    UpdateStatus picStatus, 20
    Call Write_ErrLog_File("Set appAccess = CreateObject('Access.Application')")
    Set appAccess = CreateObject("Access.Application")
    
    
    
    UpdateStatus picStatus, 30
    Call Write_ErrLog_File("accessHandle = appAccess.hWndAccessApp")
    accessHandle = appAccess.hWndAccessApp
    
    
    UpdateStatus picStatus, 40
    
'   appAccess.OpenCurrentDatabase (gsrptpath)
    Call Write_ErrLog_File("appAccess.OpenCurrentDatabase" & wsRptPath & gsDBName)
    
    appAccess.OpenCurrentDatabase wsRptPath & gsDBName, False
    
    
    
    UpdateStatus picStatus, 50
    
    wsQuery = "RPTUSRID = '" & Set_Quote(gsUserID) & "' "
    wsQuery = wsQuery & "   AND RPTDTETIM = #" & Change_SQLDate(Me.RptDteTim) & "# "

    
    wsPara = "ODBC;DRIVER=SQL Server;SERVER=" & wsServer & ";UID=" & wsUser & ";PWD=" & wsPassword & ";DATABASE=" & wsDatabase & ";"
    
    UpdateStatus picStatus, 60
    
'   Print report from Microsoft Access
    Call Write_ErrLog_File("appAccess.Run connect_report" & Trim$(wsPara) & "," & Trim$(Me.TableID))
    
    appAccess.Run "connect_report", Trim$(wsPara), Trim$(Me.TableID)
    
    
        
                  
    Select Case inMode      'Select Print Mode
        Case 0      'Preview
            
            Call Write_ErrLog_File("appAccess.DoCmd.OpenReport " & Me.RptName & "," & wsQuery)
            
            appAccess.DoCmd.OpenReport Me.RptName, acPreview, , wsQuery
            UpdateStatus picStatus, 70
            appAccess.DoCmd.Maximize
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
            
            
            appAccess.Visible = True
            
          '  Call ShowAccess(appAccess, SW_MAXIMIZE)
                        
            Call SetWindowPos(accessHandle, HWND_TOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
        
        Case 1      'Print only
        
            Call Write_ErrLog_File("appAccess.DoCmd.OpenReport " & Me.RptName & "," & wsQuery)
            appAccess.DoCmd.OpenReport Me.RptName, acNormal, , wsQuery
            UpdateStatus picStatus, 70
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
        
    End Select
              
    MousePointer = vbDefault
    
   'This is here to stop the code so you can see the report...hit F5 to
   ' continue
'   Debug.Assert 1 = 3

 '   appAccess.Quit acExit
    Set appAccess = Nothing
    
    Exit Sub
    
Access_Print_Err:

    Set appAccess = Nothing
    MsgBox "Error in Access Print! " & Err.Description
    Call Write_ErrLog_File(Err.Description)
     
    Exit Sub
    
End Sub
Private Sub Access_RunTimePrint(inMode As Integer)
    ' Declare Database variable and Microsoft Access Form variable.
    Dim wsPara As String
'    Dim appAccess As Access.Application
    Dim appAccess As Object
    Dim tmpString As String
    Dim accessHandle As Long
    Dim x As Long
    
    On Error GoTo Access_RunTimePrint_Err
    
    UpdateStatus picStatus, 10
    
    ' Return instance of Application object.
    MousePointer = vbHourglass
        
    
    UpdateStatus picStatus, 20
    
    x = Shell(gsRTPath & "msaccess.exe " & _
    Chr$(34) & wsRptPath & gsDBName & Chr$(34) & _
    "/Runtime /Wrkgrp " & Chr$(34) & _
    gsRTPath & "system.mdw" & Chr$(34), vbMaximizedFocus)
    
    UpdateStatus picStatus, 30
    Call Write_ErrLog_File("Set appAccess = GetObject(" & wsRptPath & gsDBName & ")")
    Set appAccess = GetObject(wsRptPath & gsDBName)
    UpdateStatus picStatus, 40
    
    
    Call Write_ErrLog_File("accessHandle = appAccess.hWndAccessApp")
    accessHandle = appAccess.hWndAccessApp
    UpdateStatus picStatus, 50
    

    'appAccess.OpenCurrentDatabase wsRptPath & gsDBName, False
     'UpdateStatus picStatus, 50
    
    wsQuery = "RPTUSRID = '" & Set_Quote(gsUserID) & "' "
    wsQuery = wsQuery & "   AND RPTDTETIM = #" & Change_SQLDate(Me.RptDteTim) & "# "

    
    wsPara = "ODBC;DRIVER=SQL Server;SERVER=" & wsServer & ";UID=" & wsUser & ";PWD=" & wsPassword & ";DATABASE=" & wsDatabase & ";"
    
    UpdateStatus picStatus, 60
    
    ' Print report from Microsoft Access
    Call Write_ErrLog_File("appAccess.Run connect_report" & Trim$(wsPara) & "," & Trim$(Me.TableID))
    
    appAccess.Run "connect_report", Trim$(wsPara), Trim$(Me.TableID)
    
     
                  
    Select Case inMode      'Select Print Mode
        Case 0      'Preview
        
            Call Write_ErrLog_File("appAccess.DoCmd.OpenReport " & Me.RptName & "," & wsQuery)
            
            appAccess.DoCmd.OpenReport Me.RptName, acPreview, , wsQuery
            UpdateStatus picStatus, 70
            appAccess.DoCmd.Maximize
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
            
            
          '  appAccess.Visible = True
            Call SetWindowPos(accessHandle, HWND_TOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
        
        Case 1      'Print only
        
            
            Call Write_ErrLog_File("appAccess.DoCmd.OpenReport " & Me.RptName & "," & wsQuery)
            
            appAccess.DoCmd.OpenReport Me.RptName, acNormal, , wsQuery
            UpdateStatus picStatus, 70
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
        
    End Select
    
    
   ' Debug.Assert 1 = 3

    'appAccess.Quit acExit
    Set appAccess = Nothing
    
    MousePointer = vbDefault
    
    Exit Sub
    
Access_RunTimePrint_Err:
    Set appAccess = Nothing
    
    Call Write_ErrLog_File(Err.Description)
    
    MsgBox "Error in Access Print! " & Err.Description
    Exit Sub
    
End Sub


