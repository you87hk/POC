VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGL002 
   BorderStyle     =   1  '單線固定
   Caption         =   "Stock Transfer"
   ClientHeight    =   8565
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11910
   Icon            =   "frmGL002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   StartUpPosition =   2  '螢幕中央
   WindowState     =   2  '最大化
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11160
      OleObjectBlob   =   "frmGL002.frx":030A
      TabIndex        =   11
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
         Picture         =   "frmGL002.frx":2A0D
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnSelect 
         Caption         =   "Unselect All"
         Height          =   735
         Left            =   120
         Picture         =   "frmGL002.frx":2D17
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   10080
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   735
         Left            =   120
         Picture         =   "frmGL002.frx":3021
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboBankAcc 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   1812
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Selection Criteria"
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   11535
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   6660
         Begin VB.OptionButton optBy 
            Caption         =   "BARCODE"
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   4
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton optBy 
            Caption         =   "CALLNO"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin MSMask.MaskEdBox medDocFr 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medDocTo 
         Height          =   285
         Left            =   5520
         TabIndex        =   2
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medChqDate 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDspBankAcc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   4200
         TabIndex        =   21
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblChqDate 
         Caption         =   "DOCDATE"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1740
         Width           =   1920
      End
      Begin VB.Label lblDocTo 
         Caption         =   "DOCDATE"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label lblDocFr 
         Caption         =   "DOCFR"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   780
         Width           =   1920
      End
      Begin VB.Label lblBankAcc 
         Caption         =   "BANKACC"
         Height          =   225
         Left            =   480
         TabIndex        =   14
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
      TabIndex        =   12
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
            Picture         =   "frmGL002.frx":332B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":3645
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":3A97
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":3DB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":40CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":4527
            Key             =   "book"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":4843
            Key             =   "book1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":4B63
            Key             =   "StockIn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGL002.frx":4E87
            Key             =   "StockOut"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   10
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
      Height          =   5610
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   9895
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
End
Attribute VB_Name = "frmGL002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private wsConnTime As String
Dim wsFormID As String
Private Const wsKeyType = "BNKREC"

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



Private Const GTYPE = 0
Private Const GCHQID = 1
Private Const GCHQNO = 2
Private Const GCHQDATE = 3
Private Const GCURR = 4
Private Const GCHQAMT = 5
Private Const GSTATUS = 6
Private Const GDATECLEAN = 7




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

' Call UnLockAll(wsConnTime, wsFormID)
 Ini_Scr

End Sub





Private Sub Ini_Scr()

   Me.Caption = wsFormCaption
  'lblSummary.Caption = ""
   
   
   Call SetDateMask(medDocFr)
   Call SetDateMask(medDocTo)
   Call SetDateMask(medChqDate)
   
   cboBankAcc.Text = ""
   lblDspBankAcc.Caption = ""
   medDocFr.Text = Left(gsSystemDate, 8) & "01"
   medDocTo.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medDocFr.Text))), "YYYY/MM/DD")
   medChqDate.Text = gsSystemDate
   optBy(0).Value = True
   
   UpdStatusBar picStatus, 0
   
   IniColHeader
   
  'DoEvents
   

End Sub




Private Sub cmdDelete_Click()
    
    Call cmdSave
    
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
     
          Ini_Scr_AfrKey
             
        Case vbKeyF11
     
             cmdCancel
        
        Case vbKeyF5
     
          
             
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


   Call UnLockAll(wsConnTime, wsFormID)
   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmGL002 = Nothing

End Sub


Private Sub IniColHeader()

   Dim clmX As columnheader
   
   On Error GoTo IniColHeader_Err
   
   lstData.ListItems.Clear
   lstData.ColumnHeaders.Clear
   
   

    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GTYPE"), 1000)
         clmX.Alignment = lvwColumnLeft
    Set clmX = lstData.ColumnHeaders. _
         Add(, , "KEY", 0)
    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GCHQNO"), 2000)
         clmX.Alignment = lvwColumnRight
    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GCHQDATE"), 1500)
         clmX.Alignment = lvwColumnCenter
        Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GCURR"), 1000)
         clmX.Alignment = lvwColumnCenter
    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GCHQAMT"), 2000)
         clmX.Alignment = lvwColumnRight
    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GSTATUS"), 1000)
         clmX.Alignment = lvwColumnCenter
    Set clmX = lstData.ColumnHeaders. _
         Add(, , Get_Caption(waScrItm, "GDATECLEAN"), 1000)
         clmX.Alignment = lvwColumnCenter
            
   
   
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





Private Function LoadRecord() As Boolean

   Dim rsRcd As New ADODB.Recordset
   Dim wsSQL As String
   Dim itmX As ListItem
   Dim subX As ListSubItem
   Dim wiCol As Integer
   Dim NoOfRecord As Long
   Dim wiStatus As Integer
       
    Me.MousePointer = vbHourglass
    
    LoadRecord = False

    If InputValidation = False Then
    Me.MousePointer = vbNormal
    Exit Function
    End If

    lstData.ListItems.Clear

    'Create ARCHEQUE sql
    wsSQL = " SELECT ARCQCHQID, ARCQCHQNO, ARCQCHQDATE, ARCQCURR, ARCQCHQAMT, ARCQBNKFLG, ARCQBNKDATE "
    wsSQL = wsSQL & " FROM ARCHEQUE "
    wsSQL = wsSQL & " WHERE ARCQCHQDATE BETWEEN '" & Set_Quote(medDocFr.Text) & "' AND '" & Set_Quote(medDocTo.Text) & "'"
    wsSQL = wsSQL & " AND ARCQSTATUS <> '2' "
    wsSQL = wsSQL & " AND ARCQBANKML = '" & Set_Quote(cboBankAcc.Text) & "' "
    
    If Opt_Getfocus(optBy, 2, 0) = 1 Then
    wsSQL = wsSQL & " AND ARCQBNKFLG <> 'U' "
    End If
    
   rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
   If rsRcd.RecordCount > 0 Then
        NoOfRecord = rsRcd.RecordCount
        rsRcd.MoveFirst
         Do While Not rsRcd.EOF
   
         Set itmX = lstData.ListItems.Add(, , "AR", , "book1")
         Set subX = itmX.ListSubItems.Add(GCHQID, , ReadRs(rsRcd, "ARCQCHQID"))
         Set subX = itmX.ListSubItems.Add(GCHQNO, , ReadRs(rsRcd, "ARCQCHQNO"))
         Set subX = itmX.ListSubItems.Add(GCHQDATE, , ReadRs(rsRcd, "ARCQCHQDATE"))
         Set subX = itmX.ListSubItems.Add(GCURR, , ReadRs(rsRcd, "ARCQCURR"))
         Set subX = itmX.ListSubItems.Add(GCHQAMT, , Format(To_Value(ReadRs(rsRcd, "ARCQCHQAMT")), gsAmtFmt))
         Set subX = itmX.ListSubItems.Add(GSTATUS, , ReadRs(rsRcd, "ARCQBNKFLG"))
         Set subX = itmX.ListSubItems.Add(GDATECLEAN, , ReadRs(rsRcd, "ARCQBNKDATE"))
        
        wiStatus = wiStatus + Fix((1 / NoOfRecord) * (100))
        UpdStatusBar picStatus, wiStatus
        
        rsRcd.MoveNext
        Loop
   End If
   rsRcd.Close
   
   UpdStatusBar picStatus, 100, True
   
   'Create ARCHEQUE sql
    wsSQL = " SELECT APCQCHQID, APCQCHQNO, APCQCHQDATE, APCQCURR, APCQCHQAMT, APCQBNKFLG, APCQBNKDATE "
    wsSQL = wsSQL & " FROM APCHEQUE "
    wsSQL = wsSQL & " WHERE APCQCHQDATE BETWEEN '" & Set_Quote(medDocFr.Text) & "' AND '" & Set_Quote(medDocTo.Text) & "'"
    wsSQL = wsSQL & " AND APCQSTATUS <> '2' "
    wsSQL = wsSQL & " AND APCQBANKML = '" & Set_Quote(cboBankAcc.Text) & "' "
    
    If Opt_Getfocus(optBy, 2, 0) = 1 Then
    wsSQL = wsSQL & " AND APCQBNKFLG <> 'U' "
    End If
    
   rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
   
  
   If rsRcd.RecordCount > 0 Then
         NoOfRecord = rsRcd.RecordCount
         rsRcd.MoveFirst
         Do While Not rsRcd.EOF
   
   
         Set itmX = lstData.ListItems.Add(, , "AP", , "book1")
          Set subX = itmX.ListSubItems.Add(GCHQID, , ReadRs(rsRcd, "APCQCHQID"))
         Set subX = itmX.ListSubItems.Add(GCHQNO, , ReadRs(rsRcd, "APCQCHQNO"))
         Set subX = itmX.ListSubItems.Add(GCHQDATE, , ReadRs(rsRcd, "APCQCHQDATE"))
         Set subX = itmX.ListSubItems.Add(GCURR, , ReadRs(rsRcd, "APCQCURR"))
         Set subX = itmX.ListSubItems.Add(GCHQAMT, , Format(To_Value(ReadRs(rsRcd, "APCQCHQAMT")), gsAmtFmt))
         Set subX = itmX.ListSubItems.Add(GSTATUS, , ReadRs(rsRcd, "APCQBNKFLG"))
         Set subX = itmX.ListSubItems.Add(GDATECLEAN, , ReadRs(rsRcd, "APCQBNKDATE"))
         
        wiStatus = wiStatus + Fix((1 / NoOfRecord) * (100))
        UpdStatusBar picStatus, wiStatus
        
        rsRcd.MoveNext
        Loop
   
   End If
   
   UpdStatusBar picStatus, 100, True
   rsRcd.Close
   LoadRecord = True
   Me.MousePointer = vbNormal
   Set rsRcd = Nothing
   Set itmX = Nothing
   Set subX = Nothing


   Exit Function
   
LoadRecord_Err:
   MsgBox Err.Description
   On Error Resume Next
   Me.MousePointer = vbNormal
   rsRcd.Close
   Set rsRcd = Nothing
   Set itmX = Nothing
   Set subX = Nothing

End Function


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
                                     
         

      Case Else
         .Sorted = False
         
      
      End Select
      
      
    '  lblSummary.Caption = columnheader.Text & " : " & wxSummary(1, columnheader.Index)
   End With
   
   MousePointer = vbDefault
   lstData.MousePointer = ccDefault

End Sub




Private Sub Ini_Caption()
   
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblBankAcc.Caption = Get_Caption(waScrItm, "BANKACC")
    lblDocFr.Caption = Get_Caption(waScrItm, "DOCFR")
    lblDocTo.Caption = Get_Caption(waScrItm, "DOCTO")
    optBy(0).Caption = Get_Caption(waScrItm, "OPT0")
    optBy(1).Caption = Get_Caption(waScrItm, "OPT1")
    lblChqDate.Caption = Get_Caption(waScrItm, "CHQDATE")
    
 
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





Private Sub medDocFr_GotFocus()
    FocusMe medDocFr
End Sub


Private Sub medDocFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medDocFr = False Then
            Exit Sub
        End If
        
        medDocTo.SetFocus
    End If
End Sub

Private Function chk_medDocFr() As Boolean

    chk_medDocFr = False
    
    If Trim(medDocFr.Text) = "/  /" Then
        gsMsg = "Must input Date!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocFr.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medDocFr) = False Then
        gsMsg = "Invalid Date Format!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocFr.SetFocus
        Exit Function
    End If
    
    
    chk_medDocFr = True
    
End Function
Private Sub medDocFr_LostFocus()
    FocusMe medDocFr, True
End Sub

Private Sub medDocTo_GotFocus()
    FocusMe medDocTo
End Sub


Private Sub medDocTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medDocTo = False Then
            Exit Sub
        End If
        
        Call Opt_Setfocus(optBy, 2, 0)
        
    End If
End Sub

Private Function chk_medDocTo() As Boolean

    chk_medDocTo = False
    
    If Trim(medDocTo.Text) = "/  /" Then
        gsMsg = "Must input Date!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocTo.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medDocTo) = False Then
        gsMsg = "Invalid Date Format!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocTo.SetFocus
        Exit Function
    End If
    
    
    chk_medDocTo = True
    
End Function

Private Sub medDocTo_LostFocus()
    FocusMe medDocTo, True
End Sub


Private Sub medChqDate_GotFocus()
    FocusMe medChqDate
End Sub


Private Sub medChqDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medChqDate = False Then
            Exit Sub
        End If
        
        cboBankAcc.SetFocus
    End If
End Sub
Private Function chk_medChqDate() As Boolean

    chk_medChqDate = False
    
    If Trim(medChqDate.Text) = "/  /" Then
        gsMsg = "Must input Date!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medChqDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medChqDate) = False Then
        gsMsg = "Invalid Date Format!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medChqDate.SetFocus
        Exit Function
    End If
    
    
    chk_medChqDate = True
    
End Function


Private Sub medChqDate_LostFocus()
    FocusMe medChqDate, True
End Sub






Private Sub optBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
       Ini_Scr_AfrKey
        medChqDate.SetFocus
        
    End If
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        
       Case tcGo
            Ini_Scr_AfrKey
            
       Case tcRefresh
          
        
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
    wsFormID = "GL002"
    wsConnTime = Dsp_Date(Now, True)
   
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







Private Function chk_cboBankAcc() As Boolean
Dim wsDesc As String
    chk_cboBankAcc = False
    
    If Trim(cboBankAcc.Text) = "" Then
        gsMsg = "Must input Bank Code!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboBankAcc.SetFocus
        Exit Function
    End If
    
    If Chk_MLClass(cboBankAcc, "B", wsDesc) = False Then
        gsMsg = "No Such Bank Account#"
        MsgBox gsMsg, vbOKOnly, gsTitle
        lblDspBankAcc.Caption = ""
        cboBankAcc.SetFocus
        Exit Function
    End If
    
    lblDspBankAcc.Caption = wsDesc
    
    chk_cboBankAcc = True
End Function
Private Sub cboBankAcc_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboBankAcc
  
    wsSQL = "SELECT MLCODE, MLDESC "
    wsSQL = wsSQL & " FROM MSTMERCHCLASS "
    wsSQL = wsSQL & " WHERE MLCODE LIKE '%" & IIf(cboBankAcc.SelLength > 0, "", Set_Quote(cboBankAcc.Text)) & "%' "
    wsSQL = wsSQL & " AND MLSTATUS <> '2' "
    wsSQL = wsSQL & " AND MLTYPE = 'B' "
    Call Ini_Combo(2, wsSQL, cboBankAcc.Left, cboBankAcc.Top + cboBankAcc.Height, tblCommon, wsFormID, "TBLBANKACC", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboBankAcc_GotFocus()
    FocusMe cboBankAcc
    Set wcCombo = cboBankAcc
End Sub

Private Sub cboBankAcc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboBankAcc, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboBankAcc = False Then Exit Sub
        
        Ini_Scr_AfrKey
        medDocFr.SetFocus
    End If
End Sub

Private Sub cboBankAcc_LostFocus()
    FocusMe cboBankAcc, True
End Sub




Private Function InputValidation() As Boolean

    InputValidation = False
    
        If chk_cboBankAcc = False Then
            Exit Function
        End If
        
        If chk_medDocFr = False Then
            Exit Function
        End If
        
        If chk_medDocTo = False Then
            Exit Function
        End If
        
      
    
    
    InputValidation = True
   
End Function

Private Function cmdSave() As Boolean

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wiRow As Integer
    Dim wiCnt As Integer
    
    On Error GoTo cmdSave_Err
    
    wiRow = 0
    If lstData.ListItems.Count <= 0 Then
        gsMsg = "No Record for process!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboBankAcc.SetFocus
        Exit Function
    End If
    
    For wiCnt = 1 To lstData.ListItems.Count
        If lstData.ListItems(wiCnt).Checked = True Then
        wiRow = wiRow + 1
        End If
    Next wiCnt
    
    If wiRow <= 0 Then
        gsMsg = "No Record for process!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboBankAcc.SetFocus
        Exit Function
    End If
    
 '    If ReadOnlyMode(wsConnTime, wsKeyType, cboBankAcc.Text, wsFormID) Then
 '           gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
 '           MsgBox gsMsg, vbOKOnly, gsTitle
 '           Exit Function
 '   End If
    
    MousePointer = vbHourglass
    
   
    '' Last Check when Add
      If chk_medChqDate = False Then
            MousePointer = vbNormal
            Exit Function
       End If
        
    
   cnCon.BeginTrans
     
   With lstData
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_GL002"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    For wiCnt = 1 To .ListItems.Count
    If .ListItems(wiCnt).Checked = True Then
    
    Call SetSPPara(adcmdSave, 1, .ListItems(wiCnt).Text)
    Call SetSPPara(adcmdSave, 2, .ListItems(wiCnt).SubItems(GCHQID))
    Call SetSPPara(adcmdSave, 3, .ListItems(wiCnt).SubItems(GCHQNO))
    Call SetSPPara(adcmdSave, 4, IIf(.ListItems(wiCnt).SubItems(GSTATUS) = "U", "R", "U"))
    Call SetSPPara(adcmdSave, 5, IIf(.ListItems(wiCnt).SubItems(GSTATUS) = "U", medChqDate.Text, ""))
    Call SetSPPara(adcmdSave, 6, IIf(wiCnt = wiRow, "Y", "N"))
    Call SetSPPara(adcmdSave, 7, wsFormID)
    Call SetSPPara(adcmdSave, 8, gsUserID)
    Call SetSPPara(adcmdSave, 9, Change_SQLDate(Now))
    adcmdSave.Execute
    
    End If
    Next wiCnt
    
   End With
    
    cnCon.CommitTrans
    
    gsMsg = "Bank Record Updated!"
    MsgBox gsMsg, vbOKOnly, gsTitle
        
    'Call UnLockAll(wsConnTime, wsFormID)
    Set adcmdSave = Nothing
    cmdSave = True
    
    MousePointer = vbDefault
    
    Call cmdCancel
    
    Exit Function
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Function

Private Sub Ini_Scr_AfrKey()
Dim wsUsrId As String
    
    
    If LoadRecord() = True Then
    '    If RowLock(wsConnTime, wsKeyType, cboBankAcc.Text, wsFormID, wsUsrId) = False Then
    '        gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
    '        MsgBox gsMsg, vbOKOnly, gsTitle
    '    End If
        cmdDelete.SetFocus
        
    End If
    
   
        
End Sub
