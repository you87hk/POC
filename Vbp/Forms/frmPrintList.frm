VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintList 
   BorderStyle     =   1  '單線固定
   Caption         =   "Print List"
   ClientHeight    =   8220
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11775
   Icon            =   "frmPrintList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11775
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog cdFont 
      Left            =   11280
      Top             =   840
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstData 
      Height          =   6810
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   12012
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
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   3120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintList.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintList.frx":2014
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
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font (F9)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblRptTitle 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11535
   End
   Begin VB.Label lblSummary 
      Caption         =   "Label1"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9255
   End
End
Attribute VB_Name = "frmPrintList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PrintFrom    As Integer = 0
Private Const PrintTo      As Integer = 1
Private Const ItemField    As Integer = 1
Private Const ItemNumFlag  As Integer = 2

Dim msFields As Variant
Dim msNoOfCol As Integer
Dim msPgmId As String
Dim msQuery As String
Dim msRptTitle As String
Dim NoOfRecord As Long
Dim wsMsg1 As String
Dim wsMsg2 As String
Dim wsMsg3 As String
Dim wxSummary As New XArrayDB
Dim wxData As New XArrayDB
Dim waScrItm As New XArrayDB


Private Const tcFont = "Font"
Private Const tcExit = "Exit"


Property Get Fields() As Variant

   Fields = msFields
   
End Property

Property Let Fields(ByVal NewFields As Variant)

   msFields = NewFields
   
End Property

Property Get NoOfCol() As Integer

   NoOfCol = msNoOfCol

End Property

Property Let NoOfCol(ByVal NewNoOfCol As Integer)

   msNoOfCol = NewNoOfCol

End Property

Property Get RptTitle() As String

   RptTitle = msRptTitle
   
End Property

Property Let RptTitle(ByVal NewRptTitle As String)

   msRptTitle = NewRptTitle
   
End Property

Property Get Query() As String

   Query = msQuery
   
End Property

Property Let Query(ByVal NewQuery As String)

   msQuery = NewQuery
   
End Property



Private Sub cmdFont()

   Dim wfFont As Font

   On Error GoTo FontErr
   
   cdFont.ShowFont
   lstData.Font.Name = cdFont.FontName
   lstData.Font.Bold = cdFont.FontBold
   lstData.Font.Italic = cdFont.FontItalic
   lstData.Font.Size = cdFont.FontSize
   lstData.Refresh
   DoEvents
   Exit Sub
   
FontErr:
   If cdFont.CancelError = True Then
      Exit Sub
   End If

End Sub

Private Sub Form_Activate()

   Me.MousePointer = vbHourglass

   Ini_Caption

   Ini_Scr

   Me.MousePointer = vbDefault

End Sub

Private Sub Ini_Scr()

   'Me.Width = Screen.Width
   'Me.Height = Screen.Height
   lblRptTitle.Caption = Me.RptTitle
   lstData.Width = Me.Width - 240
   lstData.Height = Me.Height - 2000
   lblRptTitle.Width = Me.Width - 240
   lblRptTitle.Left = 120
  ' lblRptTitle.Top = 120
   lstData.Left = 120
  ' lstData.Top = 240 + lblRptTitle.Height
'   lblSummary.Top = lstData.Top + lstData.Height + 120
'   cmdFont.Top = lblSummary.Top
'   cmdExit.Top = lblSummary.Top

   lstData.ListItems.Clear
   
   Dim lStyle As Long
   lStyle = SendMessage(lstData.hwnd, _
      LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   
   lStyle = LVS_EX_FULLROWSELECT
   Call SendMessage(lstData.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
      0, ByVal lStyle)
         
   
   
   
   lblSummary.Caption = ""
   DoEvents
   
   With cdFont
      .flags = cdlCFBoth Or cdlCFANSIOnly
      .CancelError = True
   End With
   
   IniColHeader
   'DoEvents
   
   LoadData

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode

        
        Case vbKeyF9
     
             cmdFont
             
        Case vbKeyF12
        
              Unload Me
            
    End Select
    
    
   ' KeyCode = vbDefault
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set wxSummary = Nothing
   Set wxData = Nothing
   Set waScrItm = Nothing
   Set frmPrintList = Nothing

End Sub

Private Sub IniColHeader()

   Dim wsSQL As String
   Dim wiCtr As Integer
   Dim clmX As columnheader
   Dim ColWidth As Integer
   
   ColWidth = IIf(Me.NoOfCol > 10, lstData.Width / 10, lstData.Width / Me.NoOfCol)
   For wiCtr = 1 To Me.NoOfCol
      Set clmX = lstData.ColumnHeaders. _
         Add(, , Me.Fields(wiCtr, 1), ColWidth)
      If Me.Fields(wiCtr, 2) = "N" And wiCtr > 1 Then
         clmX.Alignment = lvwColumnRight
      Else
         clmX.Alignment = lvwColumnLeft
      End If
      clmX.Tag = Me.Fields(wiCtr, 2)
   Next
            
   With lstData
      .DragMode = 0
      .Sorted = False
   End With

   Set clmX = Nothing

End Sub

Private Sub LoadData()

   Dim wiCtr As Integer
   Dim wsSQL As String
   Dim wsText As String
   Dim inpParent As Variant
   Dim wsDate As String
   Dim i As Long
   Dim wsMid As String
   Dim wiRow As Long
   Dim adReport As New ADODB.Recordset

   adReport.Open Me.Query, cnCon, adOpenStatic, adLockOptimistic
   
   If adReport.RecordCount = 0 Then
      adReport.Close
      Set adReport = Nothing
      Exit Sub
   Else
      NoOfRecord = adReport.RecordCount
      wxSummary.ReDim 1, 2, 1, Me.NoOfCol
      wxData.ReDim 1, NoOfRecord, 1, Me.NoOfCol
   End If
   
   With lstData
      For wiCtr = 1 To Me.NoOfCol
         Select Case .ColumnHeaders(wiCtr).Tag
         Case "D", "T", "C"
            wxSummary(1, wiCtr) = NoOfRecord
         Case Else
            wxSummary(1, wiCtr) = 0
         End Select
         wxSummary(2, wiCtr) = "DESC"
      Next
      wiRow = 1
      Do Until adReport.EOF
         For wiCtr = 1 To Me.NoOfCol
            Select Case .ColumnHeaders(wiCtr).Tag
            Case "N"       'NUMBER FIELD
               'inpParent = adReport(wiCtr - 1).Value
               wxSummary(1, wiCtr) = To_Value(wxSummary(1, wiCtr)) + To_Value(ReadRs(adReport, wiCtr - 1))
               wxData(wiRow, wiCtr) = To_Value(ReadRs(adReport, wiCtr - 1))
            Case "T"       'TEXT FIELD
               inpParent = Trim(adReport(wiCtr - 1).GetChunk(2048))
               wsText = ""
               If IsNull(inpParent) = False Then
                   For i = 1 To Len(inpParent)
                       wsMid = Mid(inpParent, i, 1)
                       If wsMid = Chr(13) Then
                           wsText = wsText & " "
                       Else
                           wsText = wsText & wsMid
                       End If
                   Next i
               End If
               wxData(wiRow, wiCtr) = wsText
            Case "D"
               'inpParent = adReport(wiCtr - 1).Value
               'If IsNull(inpParent) Then
               '   wsDate = ""
               'Else
               '   wsDate = inpParent
               '   wsDate = Dsp_Date(wsDate)
               'End If
               wxData(wiRow, wiCtr) = Dsp_Date(ReadRs(adReport, wiCtr - 1), , True)
            Case "C"
               'inpParent = adReport(wiCtr - 1).Value
               wxData(wiRow, wiCtr) = ReadRs(adReport, wiCtr - 1)
            End Select
         Next
         wiRow = wiRow + 1
         If wiRow Mod 500 = 0 Then
            .Refresh
            lblSummary.Caption = wsMsg1 & CStr(wiRow)
            DoEvents
         End If
         adReport.MoveNext
      Loop
   End With
   Me.MousePointer = vbDefault
   RefreshListView
   
   adReport.Close
   Set adReport = Nothing

   Exit Sub
   
LoadData_Err:
   MsgBox Err.Description
   On Error Resume Next
   adReport.Close
   Set adReport = Nothing

End Sub

Private Sub lstData_BeforeLabelEdit(Cancel As Integer)

    Cancel = True
    
End Sub

Private Sub lstData_ColumnClick(ByVal columnheader As MSComctlLib.columnheader)
   
   Dim wiSortIdx As Integer
   Dim wlItem As Long
   Dim strName As String
   Dim dDate As Date

   MousePointer = vbHourglass
   lstData.MousePointer = ccHourglass
   'DoEvents

   wiSortIdx = columnheader.Index - 1
   With lstData
      Select Case columnheader.Tag
      Case "C", "T"
         .SortKey = wiSortIdx
   
         'If wiSortIdx = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
               .SortOrder = lvwDescending
            Else
               .SortOrder = lvwAscending
            End If
         'End If
   
         wiSortIdx = columnheader.Index - 1
         .Sorted = True
      Case "D"
         .Sorted = False       'User clicked on the Date header
                                     'Use our sort routine to sort
                                     'by date
         'SendMessage lstData.hWnd, LVM_SORTITEMS, lstData.hWnd, _
            AddressOf CompareDates
         'lstData.Refresh
                                     
         'For wlItem = 0 To lstData.ListItems.Count - 1
         '   ListView_GetListItem wlItem, lstData.hWnd, strName, dDate, wiSortIdx + 1
         'Next
                                     
         If wxSummary(2, wiSortIdx + 1) = "DESC" Then
            wxData.QuickSort 1, NoOfRecord, wiSortIdx + 1, XORDER_ASCEND, XTYPE_DATE
            wxSummary(2, wiSortIdx + 1) = "ASC"
         Else
            wxData.QuickSort 1, NoOfRecord, wiSortIdx + 1, XORDER_DESCEND, XTYPE_DATE
            wxSummary(2, wiSortIdx + 1) = "DESC"
         End If
         RefreshListView

      Case Else
         .Sorted = False
         If wxSummary(2, wiSortIdx + 1) = "DESC" Then
            wxData.QuickSort 1, NoOfRecord, wiSortIdx + 1, XORDER_ASCEND, XTYPE_DOUBLE
            wxSummary(2, wiSortIdx + 1) = "ASC"
         Else
            wxData.QuickSort 1, NoOfRecord, wiSortIdx + 1, XORDER_DESCEND, XTYPE_DOUBLE
            wxSummary(2, wiSortIdx + 1) = "DESC"
         End If
         RefreshListView
      
      End Select
      
      
      lblSummary.Caption = columnheader.Text & " : " & wxSummary(1, columnheader.Index)
   End With
   
   MousePointer = vbDefault
   lstData.MousePointer = ccDefault

End Sub


Private Sub RefreshListView()

   Dim wiRow As Long
   Dim wiCol As Integer
   Dim itmX As ListItem
   Dim subX As ListSubItem
   
   With lstData
      .ListItems.Clear
      For wiRow = 1 To NoOfRecord
         For wiCol = 1 To Me.NoOfCol
            If wiCol = 1 Then
               Set itmX = .ListItems.Add(, , wxData(wiRow, wiCol))
            Else
               Set subX = itmX.ListSubItems.Add(wiCol - 1, , wxData(wiRow, wiCol))
            End If
         Next
         If wiRow Mod 500 = 0 Then
            .Refresh
            lblSummary.Caption = wsMsg2 & CStr(wiRow)
            DoEvents
         End If
      Next
   End With
   lblSummary.Caption = wsMsg3
   Set itmX = Nothing
   Set subX = Nothing

End Sub

Private Sub Ini_Caption()
   
   Call Get_Scr_Item("PRINTLIST", waScrItm)
   
   Me.Caption = Get_Caption(waScrItm, "SCRHDR")
 '  cmdFont.Caption = Get_Caption(waScrItm, "FONT")
 '  cmdExit.Caption = Get_Caption(waScrItm, "EXIT")
 '  wsMsg1 = Get_Caption(waScrItm, "PROCESS1")
 '  wsMsg2 = Get_Caption(waScrItm, "PROCESS2")
 '  wsMsg3 = Get_Caption(waScrItm, "PROCESS3")
   wsMsg1 = "1"
   wsMsg2 = "2"
   wsMsg3 = Get_Caption(waScrItm, "MSG3")

End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        Case tcFont
            cmdFont
            
        Case tcExit
            
            Unload Me
    End Select
    
End Sub
