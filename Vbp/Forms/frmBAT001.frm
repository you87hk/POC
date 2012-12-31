VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBAT001 
   BorderStyle     =   1  '單線固定
   Caption         =   "Batch Maintenance"
   ClientHeight    =   8565
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11910
   Icon            =   "frmBAT001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   StartUpPosition =   2  '螢幕中央
   WindowState     =   2  '最大化
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11160
      OleObjectBlob   =   "frmBAT001.frx":030A
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   10080
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   735
         Left            =   120
         Picture         =   "frmBAT001.frx":2A0D
         Style           =   1  '圖片外觀
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnSelect 
         Caption         =   "Unselect All"
         Height          =   735
         Left            =   120
         Picture         =   "frmBAT001.frx":2D17
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   10080
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   735
         Left            =   120
         Picture         =   "frmBAT001.frx":3021
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboBatchNoTo 
      Height          =   300
      Left            =   5310
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   1812
   End
   Begin VB.ComboBox cboBatchNoFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   1812
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Selection Criteria"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   11535
      Begin MSMask.MaskEdBox medPrdTo 
         Height          =   285
         Left            =   5205
         TabIndex        =   3
         Top             =   810
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
         Left            =   2400
         TabIndex        =   2
         Top             =   810
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
         Height          =   225
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label lblPrdTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4830
         TabIndex        =   12
         Top             =   870
         Width           =   375
      End
      Begin VB.Label lblBatchNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4830
         TabIndex        =   11
         Top             =   390
         Width           =   375
      End
      Begin VB.Label lblBatchNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   480
         TabIndex        =   10
         Top             =   390
         Width           =   1890
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
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   11505
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8280
      Width           =   11565
   End
   Begin MSComDlg.CommonDialog cdFont 
      Left            =   11280
      Top             =   840
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":3463
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":377D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":3BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":3EE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":4203
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":465F
            Key             =   "book"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":497B
            Key             =   "book1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":4C9B
            Key             =   "StockIn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBAT001.frx":4FBF
            Key             =   "StockOut"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go (F9)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lstData 
      Height          =   5850
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   10319
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblSummary 
      BorderStyle     =   1  '單線固定
      Caption         =   "Label1"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   11775
   End
End
Attribute VB_Name = "frmBAT001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim NoOfRecord As Long
Dim wxSummary As New XArrayDB
Dim wxData As New XArrayDB
Dim wsField As New XArrayDB
Dim NoOfCol As Integer


Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wcCombo As Control

Dim wsMsg1 As String
Dim wsMsg2 As String
Dim wsMsg3 As String

Private wsFormCaption As String
Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcRefresh = "Refresh"
Private Const tcFont = "Font"
Private Const tcExit = "Exit"




Private Sub cmdFont()

   Dim wfFont As Font

   On Error GoTo FontErr
   
   cdFont.ShowFont
   With lstData
        .Font.Name = cdFont.FontName
        .Font.Bold = cdFont.FontBold
        .Font.Italic = cdFont.FontItalic
        .Font.Size = cdFont.FontSize
        .Refresh
   End With
   


   
   DoEvents
   
   
   
   Exit Sub
   
FontErr:
   If cdFont.CancelError = True Then
      Exit Sub
   End If

End Sub



Private Sub cmdCancel()

   
Ini_Scr

End Sub





Private Sub Ini_Scr()

   Me.Caption = wsFormCaption
   lblSummary.Caption = ""
   
   cboBatchNoFr.Text = ""
   cboBatchNoTo.Text = ""
   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)
   
   UpdStatusBar picStatus, 0
   
   IniColHeader
   LoadRecord
   
  'DoEvents
   

End Sub



Private Sub cboBatchNoTo_LostFocus()
    FocusMe cboBatchNoTo, True
End Sub




Private Sub cmdDelete_Click()
   Dim i As Integer
   If lstData.ListItems.Count > 0 Then
    With lstData
    For i = 1 To .ListItems.Count
       If .ListItems(i).Checked = True Then
        Call cmdSave(.ListItems(i).Text, 3)
       End If
    Next i
    End With
   End If
   
    
   Call cmdCancel
   
End Sub

Private Sub cmdSelectAll_Click()
   Dim i As Integer
   If lstData.ListItems.Count > 0 Then
    With lstData
    For i = 1 To .ListItems.Count
        .ListItems(i).Checked = True
    Next i
    End With
   End If
End Sub

Private Sub cmdUnSelect_Click()
Dim i As Integer
   
   If lstData.ListItems.Count > 0 Then
    With lstData
    For i = 1 To .ListItems.Count
        .ListItems(i).Checked = False
    Next i
    End With
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode

        
        

        Case vbKeyF9
     
           LoadRecord
             
        Case vbKeyF11
     
             cmdCancel
        
        Case vbKeyF5
     
             RefreshListView
             
        Case vbKeyF6
     
             cmdFont
        
        Case vbKeyF12
        
              Unload Me
            
    End Select
    
    
   ' KeyCode = vbDefault
End Sub



Private Sub Form_Load()

   Me.MousePointer = vbHourglass

   Ini_Form

   Ini_Caption
  
   Ini_Scr

   Me.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set wxSummary = Nothing
   Set wxData = Nothing
   Set wsField = Nothing
  
   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmBAT001 = Nothing

End Sub

Private Sub LoadField()

  Dim wsSql As String
  Dim rsRcd As New ADODB.Recordset


   
   On Error GoTo LoadField_Err
   
   wsSql = " SELECT ScrFldID, ScrFldName, "
   wsSql = wsSql & " CASE WHEN USERTYPE IN (5, 6, 7, 8, 10, 11, 21, 24) THEN 'N' "
   wsSql = wsSql & " WHEN USERTYPE IN (12, 22, 80) THEN 'D' "
   wsSql = wsSql & " WHEN USERTYPE IN (19) THEN 'T' "
   wsSql = wsSql & " ELSE 'C' END AS ScrFldType FROM sysScrCaption, SYSCOLUMNS "
   wsSql = wsSql & " WHERE ScrType = 'FIL' "
   wsSql = wsSql & " AND SYSCOLUMNS.NAME = ScrFldID "
   wsSql = wsSql & " AND ScrPgmID = '" & Set_Quote(wsFormID) & "' "
   wsSql = wsSql & " AND ScrLangID = '" & gsLangID & "' "
   wsSql = wsSql & " AND ISNULL(RTRIM(ScrFldID), '') <> '' "
   wsSql = wsSql & " ORDER BY ScrSeqNo "
   rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
   
   If rsRcd.RecordCount = 0 Then
         MsgBox "No " & wsFormID & "in System"
         GoTo LoadField_Exit
         rsRcd.Close
         Set rsRcd = Nothing
   End If
   
        wsField.ReDim 1, 0, 0, 2
       
        Do While Not rsRcd.EOF
            wsField.AppendRows
            wsField(wsField.UpperBound(1), 0) = Trim(ReadRs(rsRcd, "ScrFldID"))
            wsField(wsField.UpperBound(1), 1) = Trim(ReadRs(rsRcd, "ScrFldName"))
            wsField(wsField.UpperBound(1), 2) = Trim(ReadRs(rsRcd, "ScrFldType"))
            rsRcd.MoveNext
        Loop
   
    rsRcd.Close
    Set rsRcd = Nothing

 
Exit Sub

LoadField_Err:
   'DISPLAY ERROR FUNCTION
   MsgBox "LoadField Err!"
   
LoadField_Exit:
   On Error Resume Next
   rsRcd.Close
   Set rsRcd = Nothing



End Sub
Private Sub IniColHeader()

   Dim wiCtr As Integer
   Dim clmX As columnheader
   Dim ColWidth As Integer
   
   On Error GoTo IniColHeader_Err
   
   lstData.ListItems.Clear
   lstData.ColumnHeaders.Clear
   
   NoOfRecord = 0
   NoOfCol = wsField.UpperBound(1)
   wxSummary.ReDim 1, 2, 1, NoOfCol
   wxData.ReDim 1, 0, 1, NoOfCol
   

   ColWidth = IIf(NoOfCol > 10, lstData.Width / 10, lstData.Width / NoOfCol)
   For wiCtr = 1 To NoOfCol
      Set clmX = lstData.ColumnHeaders. _
         Add(, , wsField(wiCtr, 1), IIf(wiCtr = 1, 1500, ColWidth))
      If wsField(wiCtr, 2) = "N" Then
         clmX.Alignment = lvwColumnRight
      Else
         clmX.Alignment = lvwColumnLeft
      End If
      clmX.Tag = wsField(wiCtr, 2)
      wxSummary(1, wiCtr) = 0
      wxSummary(2, wiCtr) = "DESC"
   Next
            
   With lstData
      .DragMode = 0
      .Sorted = False
   End With

   Set clmX = Nothing
   
Exit Sub
IniColHeader_Err:
   'DISPLAY ERROR FUNCTION
   MsgBox "IniColHeader Err!"
   MsgBox Err.Description
IniColHeader_Exit:
   On Error Resume Next
   Set clmX = Nothing


End Sub





Private Sub LoadData()

   Dim wiCtr As Integer
   Dim wsSql As String
   Dim wsText As String
   Dim inpParent As Variant
   Dim wsDate As String
   Dim i As Long
   Dim wsMid As String
   Dim wiRow As Long
   Dim rsRcd As New ADODB.Recordset
   Dim wiStatus As Integer
   
    
    Me.MousePointer = vbHourglass
    


    'Create Selection Criteria
    wsSql = " SELECT STHBATCHNO, SCDDESC, STHCTLPRD, STHLASTUPD, STHLASTUPDDATE "
    wsSql = wsSql & " FROM ICSTKTRNHD, SYSCODEDESC "
    wsSql = wsSql & " WHERE STHTRNCODE = SCDCODE "
    wsSql = wsSql & " AND STHBatchNo BETWEEN " & "'" & Set_Quote(cboBatchNoFr.Text) & "'" & " AND " & "'" & IIf(Trim(cboBatchNoTo.Text) = "", String(15, "z"), Set_Quote(cboBatchNoTo.Text)) & "'"
    wsSql = wsSql & " AND STHCTLPRD BETWEEN " & "'" & IIf(Trim(medPrdFr.Text) = "/", String(6, "0"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'" & " AND " & "'" & IIf(Trim(medPrdTo.Text) = "/", String(6, "9"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSql = wsSql & " AND STHSTATUS = '1'"
    wsSql = wsSql & " AND SCDLANGID = '" & gsLangID & "'"
    
   rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
   
   If rsRcd.RecordCount = 0 Then
      rsRcd.Close
      NoOfRecord = 0
      IniColHeader
      Set rsRcd = Nothing
      Me.MousePointer = vbDefault
      Exit Sub
   Else
      NoOfRecord = rsRcd.RecordCount
      wxSummary.ReDim 1, 2, 1, NoOfCol
      wxData.ReDim 1, NoOfRecord, 1, NoOfCol
   End If
   

   
   With lstData
      For wiCtr = 1 To NoOfCol
         Select Case .ColumnHeaders(wiCtr).Tag
         Case "D", "T", "C"
            wxSummary(1, wiCtr) = NoOfRecord
         Case Else
            wxSummary(1, wiCtr) = 0
         End Select
         wxSummary(2, wiCtr) = "DESC"
      Next
      wiRow = 1
      Do Until rsRcd.EOF
         For wiCtr = 1 To NoOfCol
            Select Case .ColumnHeaders(wiCtr).Tag
            Case "N"       'NUMBER FIELD
               'inpParent = rsRcd(wiCtr - 1).Value
               wxSummary(1, wiCtr) = To_Value(wxSummary(1, wiCtr)) + To_Value(ReadRs(rsRcd, wsField(wiCtr, 0)))
               wxData(wiRow, wiCtr) = To_Value(ReadRs(rsRcd, wsField(wiCtr, 0)))
            Case "T"       'TEXT FIELD
               inpParent = Trim(rsRcd(wsField(wiCtr, 0)).GetChunk(2048))
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
               'inpParent = rsRcd(wiCtr - 1).Value
               'If IsNull(inpParent) Then
               '   wsDate = ""
               'Else
               '   wsDate = inpParent
               '   wsDate = Dsp_Date(wsDate)
               'End If
               wxData(wiRow, wiCtr) = Dsp_Date(ReadRs(rsRcd, wsField(wiCtr, 0)), , True)
            Case "C"
               'inpParent = rsRcd(wiCtr - 1).Value
               wxData(wiRow, wiCtr) = ReadRs(rsRcd, wsField(wiCtr, 0))
            End Select
         Next
         wiRow = wiRow + 1
         If wiRow Mod 500 = 0 Then
            .Refresh
            lblSummary.Caption = wsMsg1 & CStr(wiRow)
            DoEvents
         End If
         rsRcd.MoveNext
         wiStatus = wiStatus + Fix((1 / NoOfRecord) * (100))
         UpdStatusBar picStatus, wiStatus
      Loop
   End With
   
   UpdStatusBar picStatus, 100, True
   Me.MousePointer = vbDefault
   
    
   
     
   RefreshListView
   
   rsRcd.Close
   Set rsRcd = Nothing

   Exit Sub
   
LoadData_Err:
   MsgBox Err.Description
   On Error Resume Next
   rsRcd.Close
   Set rsRcd = Nothing

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
   Dim wsImage As String
   

   wsImage = "book"
   
   
   With lstData
      .ListItems.Clear
      For wiRow = 1 To NoOfRecord
         For wiCol = 1 To NoOfCol
            If wiCol = 1 Then
               Set itmX = .ListItems.Add(, , wxData(wiRow, wiCol), , wsImage)
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
   
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblBatchNoFr.Caption = Get_Caption(waScrItm, "BatchNoFR")
    lblBatchNoTo.Caption = Get_Caption(waScrItm, "BatchNoTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
 
    fraSelect.Caption = Get_Caption(waScrItm, "SELECT")
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcFont).ToolTipText = Get_Caption(waScrToolTip, tcFont) & "(F6)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    cmdDelete.Caption = Get_Caption(waScrItm, "CMDDELETE")
    cmdSelectAll.Caption = Get_Caption(waScrItm, "CMDSELECTALL")
    cmdUnSelect.Caption = Get_Caption(waScrItm, "CMDUNSELECT")

    wsMsg1 = "1"
    wsMsg2 = "2"
    wsMsg3 = Get_Caption(waScrItm, "MSG3")

End Sub



Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        
       Case tcGo
            LoadRecord
            
       Case tcRefresh
            RefreshListView
        
       Case tcCancel
            cmdCancel
            
       Case tcFont
            cmdFont
            
        Case tcExit
            
            Unload Me
    End Select
    
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    wsFormID = "BAT001"
   
   lstData.SmallIcons = iglProcess
   lstData.CheckBoxes = True
   
  ' Dim lStyle As Long
  ' lStyle = SendMessage(lstData.hwnd, _
  '    LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   
  ' lStyle = LVS_EX_FULLROWSELECT
  ' Call SendMessage(lstData.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
  '    0, ByVal lStyle)
         
   
   
   With cdFont
      .flags = cdlCFBoth Or cdlCFANSIOnly
      .CancelError = True
   End With
   
   LoadField
    

 
End Sub

Private Sub tblCommon_DblClick()
    
    wcCombo.Text = tblCommon.Columns(0).Text
    wcCombo.SetFocus
    tblCommon.Visible = False
    SendKeys "{Enter}"
    
    
End Sub

Private Sub tblCommon_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        KeyCode = vbDefault
        tblCommon.Visible = False
        wcCombo.SetFocus
        
    ElseIf KeyCode = vbKeyReturn Then
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

Private Function LoadRecord() As Boolean

    
    LoadRecord = False
    
If InputValidation = False Then Exit Function
  
    
    Call LoadData
    
  
 
 
 
 LoadRecord = True
 
End Function


Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "週期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        gsMsg = "至到日期不能大於由期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
        gsMsg = "週期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function


Private Function chk_cboBatchNoTo() As Boolean
    chk_cboBatchNoTo = False
    
    If UCase(cboBatchNoFr.Text) > UCase(cboBatchNoTo.Text) Then
        gsMsg = "至到日期不能大於由期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboBatchNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboBatchNoTo = True
End Function
Private Sub cboBatchNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboBatchNoFr
  
    wsSql = "SELECT STHBatchNo, STHLASTUPDDATE "
    wsSql = wsSql & " FROM ICSTKTRNHD "
    wsSql = wsSql & " WHERE STHBatchNo LIKE '%" & IIf(cboBatchNoTo.SelLength > 0, "", Set_Quote(cboBatchNoTo.Text)) & "%' "
    wsSql = wsSql & " AND STHSTATUS  <> '2' "
    wsSql = wsSql & " AND STHTRNCODE  = '2' "
    Call Ini_Combo(2, wsSql, cboBatchNoFr.Left, cboBatchNoFr.Top + cboBatchNoFr.Height, tblCommon, wsFormID, "TBLBatchNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboBatchNoFr_GotFocus()
    FocusMe cboBatchNoFr
    Set wcCombo = cboBatchNoFr
End Sub

Private Sub cboBatchNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboBatchNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboBatchNoFr.Text) <> "" And _
            Trim(cboBatchNoTo.Text) = "" Then
            cboBatchNoTo.Text = cboBatchNoFr.Text
        End If
        cboBatchNoTo.SetFocus
    End If
End Sub

Private Sub cboBatchNoFr_LostFocus()
    FocusMe cboBatchNoFr, True
End Sub

Private Sub cboBatchNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboBatchNoTo
  
    wsSql = "SELECT STHBatchNo, STHLASTUPDDATE "
    wsSql = wsSql & " FROM ICSTKTRNHD "
    wsSql = wsSql & " WHERE STHBatchNo LIKE '%" & IIf(cboBatchNoTo.SelLength > 0, "", Set_Quote(cboBatchNoTo.Text)) & "%' "
    wsSql = wsSql & " AND STHSTATUS  <> '2' "
    wsSql = wsSql & " AND STHTRNCODE  = '2' "
    Call Ini_Combo(2, wsSql, cboBatchNoTo.Left, cboBatchNoTo.Top + cboBatchNoTo.Height, tblCommon, wsFormID, "TBLBatchNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboBatchNoTo_GotFocus()
    FocusMe cboBatchNoTo
    Set wcCombo = cboBatchNoTo
End Sub

Private Sub cboBatchNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboBatchNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboBatchNoTo = False Then
            Exit Sub
        End If
        
        medPrdFr.SetFocus
        
    End If
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
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub
Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        Call LoadRecord
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub
Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboBatchNoTo = False Then
        cboBatchNoTo.SetFocus
        Exit Function
    End If
    
    
    If chk_medPrdTo = False Then
        medPrdTo.SetFocus
        Exit Function
    End If
    
    InputValidation = True
   
End Function

Private Function cmdSave(ByVal InDocNo As String, ByVal inAction As Integer) As Boolean

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
   
    '' Last Check when Add
    
   
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_UPDBATCH"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, inAction)
    Call SetSPPara(adcmdSave, 2, InDocNo)
    adcmdSave.Execute
    

    cnCon.CommitTrans
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Set adcmdSave = Nothing
    cmdSave = True
    
    MousePointer = vbDefault
    
    Exit Function
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Function
