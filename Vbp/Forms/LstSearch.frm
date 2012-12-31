VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLstSearch 
   BorderStyle     =   1  '單線固定
   Caption         =   "Quick Search"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   1695
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11850
   Begin VB.OptionButton optSortDesc 
      Height          =   180
      Left            =   2520
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optSortAsc 
      Height          =   180
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Frame fraFilter 
      Height          =   810
      Left            =   5950
      TabIndex        =   6
      Top             =   360
      Width           =   5820
      Begin VB.CommandButton btnReset 
         Caption         =   "&Reset"
         Height          =   280
         Left            =   4560
         TabIndex        =   10
         Top             =   160
         Width           =   1095
      End
      Begin VB.TextBox txtFilterValue 
         Height          =   270
         Left            =   2940
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cboFilterField 
         Height          =   300
         ItemData        =   "LstSearch.frx":0000
         Left            =   120
         List            =   "LstSearch.frx":0002
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filtering :"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1920
      End
   End
   Begin MSAdodcLib.Adodc OLDadoItems 
      Height          =   330
      Left            =   5160
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraFind 
      Height          =   810
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5820
      Begin VB.CommandButton btnFindFirst 
         Caption         =   "Find &First"
         Height          =   280
         Left            =   3360
         TabIndex        =   16
         Top             =   160
         Width           =   1095
      End
      Begin VB.CommandButton btnFindNext 
         Caption         =   "Find &Next"
         Height          =   280
         Left            =   4560
         TabIndex        =   15
         Top             =   160
         Width           =   1095
      End
      Begin VB.ComboBox cboSearchField 
         Height          =   300
         ItemData        =   "LstSearch.frx":0004
         Left            =   120
         List            =   "LstSearch.frx":0006
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtOutput 
         Height          =   270
         Left            =   2940
         TabIndex        =   0
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblFind 
         Caption         =   "Find :"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1920
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   5760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":0008
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":08E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":0BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":14D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":192A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":1D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":2096
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":24E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":293A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":2C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LstSearch.frx":2F6E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
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
            Key             =   "Locate"
            Object.ToolTipText     =   "Locate"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSDataGridLib.DataGrid grdItems 
      Bindings        =   "LstSearch.frx":33C0
      Height          =   5475
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   9657
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  '靠右對齊
      Caption         =   "Descending"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblAsc 
      Alignment       =   1  '靠右對齊
      Caption         =   "Ascending"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmLstSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sBindSQL As String                   '-- SQL to bind to the ADO control.
Public sBindWhereSQL As String
Public sBindOrderSQL As String

'Public sCountSQL As String                 '-- Display total no. of records SQL.
Public vHeadDataAry As Variant              '-- Two dimension arary. 1-Field description; 2-Field name.
Public vFilterAry As Variant                '-- Two dimension arary. 1-Field description; 2-Field name.
Public vSearchAry As Variant                '-- Two dimension arary. 1-Field description; 2-Field name.
Public oTextBox As TextBox                  '-- Calling form ADO control reference to locate record purpose.
Public lKeyPos As Long                      '-- For Recordset.Find purpose.
Public lSearchPos As Long                   '-- Check the type of sKeyName (String or Numeric).
Private isFinding As Boolean
Private sWhere1SQL As String

Private Sub btnFindFirst_Click()
    FindRecord "FindFirst"
End Sub

Private Sub btnFindNext_Click()
    FindRecord "FindNext"
End Sub

Private Sub btnReset_Click()
    
    Me.MousePointer = vbHourglass
    txtFilterValue = ""
    cboFilterField.ListIndex = -1
    sWhere1SQL = ""
    
    With AdoItems
        .RecordSource = sBindSQL & sBindWhereSQL & sBindOrderSQL
        .Refresh
    End With

    Format_Grid

    grdItems.SetFocus
    'SendKeys "{RIGHT}"
    'SendKeys "{LEFT}"
    Me.MousePointer = vbNormal
End Sub

Private Sub cboSearchField_Change()
    isFinding = False
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    Dim iCounter As Integer
    
    Me.WindowState = 0
    
    With AdoItems
        .ConnectionString = gsConnectString
        .RecordSource = sBindSQL & sBindWhereSQL & sBindOrderSQL
        .LockType = adLockOptimistic
        '.CursorType = adOpenDynamic
        '.CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .CursorLocation = adUseServer
        .Refresh
    End With
    
    '-- Bind datagrid caption and datafield.
    Format_Grid
    'With grdItems
    '    For icounter = 1 To UBound(vHeadDataAry)
    '        .Columns(icounter - 1).Caption = vHeadDataAry(icounter, 1)
    '        .Columns(icounter - 1).DataField = vHeadDataAry(icounter, 2)
    '        .Columns(icounter - 1).Width = vHeadDataAry(icounter, 3)
    '    Next
    'End With
    
    With cboSearchField
        For iCounter = 1 To UBound(vSearchAry)
            .AddItem vSearchAry(iCounter, 1)
        Next
    End With
    
    With cboFilterField
        For iCounter = 1 To UBound(vFilterAry)
            .AddItem vFilterAry(iCounter, 1)
        Next
    End With
End Sub

Private Sub Form_Resize()
    Dim iWidth As Integer
    
    If Me.WindowState = 0 Then
        Me.Height = 7580
        Me.Width = 11940
        Me.Left = 0
        Me.Top = 0
    End If
End Sub

Private Sub grdItems_DblClick()
    Locate_Key
End Sub

Private Sub grdItems_HeadClick(ByVal ColIndex As Integer)
    Me.MousePointer = vbHourglass
    With AdoItems
        If optSortAsc.Value = True Then
            .RecordSource = sBindSQL & sBindWhereSQL & sWhere1SQL & " ORDER BY " + grdItems.Columns(ColIndex).DataField & " ASC"
        Else
            .RecordSource = sBindSQL & sBindWhereSQL & sWhere1SQL & " ORDER BY " + grdItems.Columns(ColIndex).DataField & " DESC"
        End If
        .Refresh
    End With
    
    Format_Grid
    grdItems.SetFocus
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    Me.MousePointer = vbNormal
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Locate"
            Locate_Key
        Case "Exit"
            Unload Me
        Case "Refresh"
            AdoItems.Refresh
            grdItems.Refresh
    End Select
End Sub

Private Sub Locate_Key()
    If AdoItems.Recordset.BOF Or AdoItems.Recordset.EOF Then
        'sMsg = GetErrorMessage("PSH0003")
        sMsg = "檔案找不到, 請重新整理再選取想要之資料!"
        MsgBox sMsg, vbInformation + vbOKOnly, gsTitle
    Else
        oTextBox = grdItems.Columns(lKeyPos - 1)
        Unload Me
    End If
End Sub

Private Sub FindRecord(Optional sType)
    Dim isFound As Variant
    
    If IsMissing(sType) Then
        If txtOutput <> "" And cboSearchField.ListIndex <> -1 Then
            Me.MousePointer = vbHourglass
            If Not isFinding Then
                AdoItems.Recordset.MoveFirst
                AdoItems.Recordset.Find vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '" & txtOutput & "%'"
                isFinding = True
            Else
                AdoItems.Recordset.Find vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '" & txtOutput & "%'", 1, adSearchForward
            End If
            
            grdItems.SetFocus
            SendKeys "{RIGHT}"
            SendKeys "{LEFT}"
            Me.MousePointer = vbNormal
        End If
    Else
        If UCase(sType) = "FINDFIRST" Then
            If txtOutput <> "" And cboSearchField.ListIndex <> -1 Then
                Me.MousePointer = vbHourglass
                AdoItems.Recordset.MoveFirst
                AdoItems.Recordset.Find vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '" & txtOutput & "%'"
                isFinding = True
                
                grdItems.SetFocus
                SendKeys "{RIGHT}"
                SendKeys "{LEFT}"
                Me.MousePointer = vbNormal
            End If
        ElseIf UCase(sType) = "FINDNEXT" Then
            If txtOutput <> "" And cboSearchField.ListIndex <> -1 Then
                Me.MousePointer = vbHourglass
                    
                AdoItems.Recordset.Find vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '" & txtOutput & "%'", 1, adSearchForward
                isFinding = True
                
                grdItems.SetFocus
                SendKeys "{RIGHT}"
                SendKeys "{LEFT}"
                Me.MousePointer = vbNormal
            End If
        End If
    End If
End Sub

Private Sub txtFilterValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        Call goFilter
    End If
End Sub

Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        Call FindRecord
    Else
        isFinding = False
    End If
End Sub

Private Sub goFilter()
    If txtFilterValue <> "" And cboFilterField.ListIndex <> -1 Then
        Me.MousePointer = vbHourglass
        'adoItems.Recordset.MoveFirst
        'adoItems.Recordset.Find vFilterAry(cboSearchField.ListIndex + 1, 2) & " like '" & txtOutput & "%'"
        If InStr(1, UCase(sBindWhereSQL), "WHERE", vbTextCompare) <> 0 Then
            sWhere1SQL = " AND " + vFilterAry(cboFilterField.ListIndex + 1, 2) & " LIKE '" & txtFilterValue & "%'"
        Else
            sWhere1SQL = " WHERE " + vFilterAry(cboFilterField.ListIndex + 1, 2) & " LIKE '" & txtFilterValue & "%'"
        End If
        
        With AdoItems
            .RecordSource = sBindSQL & sBindWhereSQL & sWhere1SQL & sBindOrderSQL
            .Refresh
        End With
    
        Format_Grid
        
        grdItems.SetFocus
        SendKeys "{RIGHT}"
        SendKeys "{LEFT}"
        Me.MousePointer = vbNormal
    End If
End Sub

Private Sub Format_Grid()
    Dim iCounter As Integer
    
    With grdItems
        For iCounter = 1 To UBound(vHeadDataAry)
            .Columns(iCounter - 1).Caption = vHeadDataAry(iCounter, 1)
            .Columns(iCounter - 1).DataField = vHeadDataAry(iCounter, 2)
            .Columns(iCounter - 1).Width = vHeadDataAry(iCounter, 3)
        Next
    End With
End Sub

Private Sub tblCommon_DblClick()
    
    'If wcCombo.Name = tblDetail.Name Then
    '    tblDetail.EditActive = True
    '    Select Case wcCombo.Col
    '      Case BOOKCODE
    '           wcCombo.Text = tblCommon.Columns(0).Text
    '      Case Else
    '           wcCombo.Text = tblCommon.Columns(0).Text
    '   End Select
    'Else
    '   wcCombo.Text = tblCommon.Columns(0).Text
    'End If
    
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
        'If wcCombo.Name = tblDetail.Name Then
        '    tblDetail.EditActive = True
        '    Select Case wcCombo.Col
        '      Case BOOKCODE
        '           wcCombo.Text = tblCommon.Columns(0).Text
        '      Case Else
        '           wcCombo.Text = tblCommon.Columns(0).Text
        '   End Select
        'Else
        '   wcCombo.Text = tblCommon.Columns(0).Text
        'End If
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

