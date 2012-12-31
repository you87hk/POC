VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form nbMain 
   Caption         =   "Main Menu"
   ClientHeight    =   750
   ClientLeft      =   735
   ClientTop       =   2640
   ClientWidth     =   11880
   Icon            =   "nbMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleWidth      =   11880
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   1800
      Top             =   1440
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
            Picture         =   "nbMain.frx":030A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":0626
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":0942
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":0C5E
            Key             =   "book"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":0F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":129A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":15BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":1F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":2366
            Key             =   "bigBook"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":2686
            Key             =   "UTILITY"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":2ADE
            Key             =   "REPORT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":2DFE
            Key             =   "INQUIRY"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":311E
            Key             =   "TRANSFER"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":343E
            Key             =   "OPERATION"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":375E
            Key             =   "MASTER"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":403A
            Key             =   "ACCOUNT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":4356
            Key             =   "FILE"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":6062
            Key             =   "INVENTORY"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":693E
            Key             =   "PO"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":6D96
            Key             =   "ACCRPT"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "nbMain.frx":70BA
            Key             =   "LIST"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  '對齊表單下方
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15293
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2010/03/24"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "下午 05:43"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prev"
            Object.ToolTipText     =   "Previous Page (F2)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Page (F3)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3000
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OK"
            Object.ToolTipText     =   "OK (F10)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出系統 (F12)"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cboCommand 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   0
         Width           =   2970
      End
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   7455
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   13150
      _Version        =   393217
      Style           =   7
      ImageList       =   "iglProcess"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwDB 
      Height          =   7455
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   13150
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "iglProcess"
      SmallIcons      =   "iglProcess"
      ColHdrIcons     =   "iglProcess"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuMasterSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "&Operation"
      Begin VB.Menu mnuOperationSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuPO 
      Caption         =   "&PO"
      Begin VB.Menu mnuPOSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "In&Ventory"
      Begin VB.Menu mnuInventorySub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuAccount 
      Caption         =   "&Account"
      Begin VB.Menu mnuAccountSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuInquiry 
      Caption         =   "&Inquiry"
      Begin VB.Menu mnuInquirySub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuReportSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuACCRPT 
      Caption         =   "Acc Report"
      WindowList      =   -1  'True
      Begin VB.Menu mnuACCRPTSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuUtilitySub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Begin VB.Menu mnuListSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "nbMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mNode As Node ' Module-level variable for Nodes'
Private mItem As ListItem ' Module-level ListItem variable.
Private EventFlag As Integer ' To signal which event has occurred.
Private sCurrentIndex As String
Private mStatusBarStyle As Integer ' Switches Statusbar style

Const ROOT = 1 ' For EventFlag, Signals Publisher colmunheader objects.
Const TITLE = 2 ' EventFlag, signals Title in ListView

Private Const tcPrev = "Prev"
Private Const tcNext = "Next"
Private Const tcOK = "OK"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Dim waFileSub As New XArrayDB
Dim waMasterSub As New XArrayDB
Dim waOperationSub As New XArrayDB
Dim waPOSub As New XArrayDB
Dim waInventorySub As New XArrayDB
Dim waACCOUNTSub As New XArrayDB
Dim waInquirySub As New XArrayDB
Dim waReportSub As New XArrayDB
Dim waUtilitySub As New XArrayDB
Dim waAccRptSub As New XArrayDB
Dim waListSub As New XArrayDB



Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB

Dim wsFormID As String

Private Sub IniForm()
    Me.KeyPreview = True
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = 900
    
    wsFormID = "MAIN"

    giCurrIndex = -1
    
    With tbrMain
                .Buttons(tcPrev).Enabled = False
                .Buttons(tcNext).Enabled = False
    End With
End Sub

Private Sub ChgPrevPage()
    If giCurrIndex = 0 Then
        tbrMain.Buttons(tcPrev).Enabled = False
    End If
    
    If giCurrIndex <> -1 Then
        
        Call Call_Pgm(waFileSub, 0, UCase(cboCommand.List(giCurrIndex)), 1)
        giCurrIndex = giCurrIndex - 1
        tbrMain.Buttons(tcNext).Enabled = True
        
    End If
End Sub

Private Sub ChgNextPage()
    giCurrIndex = giCurrIndex + 1
    
    If giCurrIndex = cboCommand.ListCount - 1 Then
        tbrMain.Buttons(tcNext).Enabled = False
    End If
    
    If giCurrIndex < cboCommand.ListCount Then
        
        Call Call_Pgm(waFileSub, 0, UCase(cboCommand.List(giCurrIndex)), 1)
        tbrMain.Buttons(tcPrev).Enabled = True
        
    End If
End Sub

Private Sub Ini_Menu()
        
    ' First node with 'Root' as text.
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
        
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    mnuFile.Caption = Get_Caption(waScrItm, "FILE")
    mnuMaster.Caption = Get_Caption(waScrItm, "MASTER")
    mnuOperation.Caption = Get_Caption(waScrItm, "OPERATION")
    mnuPO.Caption = Get_Caption(waScrItm, "PO")
    mnuInventory.Caption = Get_Caption(waScrItm, "INVENTORY")
    mnuAccount.Caption = Get_Caption(waScrItm, "ACCOUNT")
    mnuInquiry.Caption = Get_Caption(waScrItm, "INQUIRY")
    mnuReport.Caption = Get_Caption(waScrItm, "REPORT")
    mnuUtility.Caption = Get_Caption(waScrItm, "UTILITY")
    mnuACCRPT.Caption = Get_Caption(waScrItm, "ACCRPT")
    mnuList.Caption = Get_Caption(waScrItm, "LIST")
    
    
    tbrMain.Buttons(tcOK).ToolTipText = Get_Caption(waScrToolTip, tcOK) & "(F10)"
    tbrMain.Buttons(tcPrev).ToolTipText = Get_Caption(waScrToolTip, tcPrev) & "(F2)"
    tbrMain.Buttons(tcNext).ToolTipText = Get_Caption(waScrToolTip, tcNext) & "(F3)"
    tbrMain.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrMain.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    Call Ini_PgmMenu(mnuFileSub, "FILE", waFileSub)
    Call Ini_PgmMenu(mnuMasterSub, "MASTER", waMasterSub)
    Call Ini_PgmMenu(mnuOperationSub, "OPERATION", waOperationSub)
    Call Ini_PgmMenu(mnuPOSub, "PO", waPOSub)
    Call Ini_PgmMenu(mnuInventorySub, "INVENTORY", waInventorySub)
    Call Ini_PgmMenu(mnuAccountSub, "ACCOUNT", waACCOUNTSub)
    Call Ini_PgmMenu(mnuInquirySub, "INQUIRY", waInquirySub)
    Call Ini_PgmMenu(mnuReportSub, "REPORT", waReportSub)
    Call Ini_PgmMenu(mnuUtilitySub, "UTILITY", waUtilitySub)
    Call Ini_PgmMenu(mnuACCRPTSub, "ACCRPT", waAccRptSub)
    Call Ini_PgmMenu(mnuListSub, "LIST", waListSub)
    
    
    
    staMain.Panels(1).Text = gsComNam
    
    lvwDB.ColumnHeaders.Clear
    lvwDB.ColumnHeaders.Add , , Get_Caption(waScrItm, "LST01"), 1500
    lvwDB.ColumnHeaders.Add , , Get_Caption(waScrItm, "LST02"), 5000
End Sub

Private Sub cboCommand_GotFocus()
    FocusMe cboCommand
End Sub

Private Sub cboCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Call_Pgm(waFileSub, 0, UCase(cboCommand.Text))
    End If
End Sub

Private Sub cboCommand_LostFocus()
    FocusMe cboCommand, True
End Sub

Private Sub Form_Activate()
If Me.WindowState = 0 Then
    If Forms.Count = 1 Then
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = 9000
    End If
End If
End Sub

Private Sub Form_Deactivate()
    
 If Me.WindowState = 0 Then
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = 1455
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        
        
        
        Case vbKeyF2
        
         
             ChgPrevPage
        
        Case vbKeyF3
             
             ChgNextPage
             
        Case vbKeyF10
        
             Call Call_Pgm(waFileSub, 0, UCase(cboCommand.Text))
            
        Case vbKeyF11
        
             Call cmdCancel
        
        Case vbKeyF12
        
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()

Call IniForm
Call Ini_Menu
Call Ini_TreeView


End Sub


Private Sub Ini_TreeView()
    
   ' lvwDB.View = lvwReport
   
    '= lvwReport
    ' Add three panels, and set Autosize for each

    Dim i As Integer
    
   Dim lStyle As Long
   lStyle = SendMessage(lvwDB.hwnd, _
      LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   
   lStyle = LVS_EX_FULLROWSELECT
   Call SendMessage(lvwDB.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
      0, ByVal lStyle)
    
    
    
    ' Configure TreeView
   ' tvwDB.Sorted = True
    Set mNode = tvwDB.Nodes.Add()
    mNode.Text = gsComNam
    mNode.Tag = Me.Name
    mNode.Image = "closed"
    tvwDB.LabelEdit = False
    EventFlag = 0
    
    
    LoadList
    ' If the Biblio database can't be found, open the
    ' common dialog control to let the user find it.

End Sub
Private Sub LoadList()
    ' Declare variables for the Data Access objects.
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    
    Dim intIndex ' Variable for index of current node.
    Dim wsTmp As String
    
        wsSQL = "select ScrPgmID , min(ScrSeqNo) Seq from sysScrCaption "
        wsSQL = wsSQL & " WHERE ScrType = 'MNU' "
      '  wsSql = wsSql & " AND ScrPgmID <> 'FILE' "
        wsSQL = wsSQL & " AND ScrPgmID <> 'POPUP' "
        wsSQL = wsSQL & " AND ScrType = 'MNU' "
        wsSQL = wsSQL & " AND ScrLangID = '" & gsLangID & "' "
        wsSQL = wsSQL & " Group By ScrPgmID "
        wsSQL = wsSQL & " Order By Seq "
        
        
        Set mNode = tvwDB.Nodes.Add(1, tvwChild, "ROOT", Get_Caption(waScrItm, "ROOT"), "closed")
        mNode.Tag = "Root" ' Identifies the table.
        intIndex = mNode.Index
       
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
        If rsRcd.RecordCount > 0 Then
        rsRcd.MoveFirst
        Do While Not rsRcd.EOF
         
            
            wsTmp = Get_Caption(waScrItm, ReadRs(rsRcd, "ScrPgmID"))
            wsTmp = SkipA(wsTmp)
            Set mNode = tvwDB.Nodes.Add(intIndex, tvwChild)
            mNode.Text = wsTmp
            mNode.Key = ReadRs(rsRcd, "ScrPgmID")  ' Unique ID.
            mNode.Tag = "Item"       ' Table name.
            mNode.Image = ReadRs(rsRcd, "ScrPgmID")     ' Image from ImageList.
        
        rsRcd.MoveNext
        Loop
        End If
            
        
     ' Sort the CardClass nodes.
  '  tvwDB.Nodes(1).Sorted = True
    ' Expand top node.
    tvwDB.Nodes(1).Expanded = True
    tvwDB.Nodes(2).Expanded = True
    rsRcd.Close
    Set rsRcd = Nothing
    
    ' configure statusbar.
   ' CardClassStatusBar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sMsg As String
    Dim sSQL As String
On Error GoTo ErrHand

    sMsg = "Are you sure to exit this system?" & Chr(10) & Chr(10)
    sMsg = sMsg & "請問你是不是肯定退出這系統?"

    If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, gsTitle) = vbNo Then
        
 
 '       sSQL = "DUMP TRANSACTION CHUNGFAIDB WITH NO_LOG"
 '       cnCon.Execute sSQL
 '
 '
 '   Else
        Cancel = True
    End If

Exit Sub

ErrHand:
     MsgBox Err.Description
     Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim f As Form
    
    For Each f In Forms
    
    If f.Name <> Me.Name Then
    MsgBox "Please Contact Nbase : Can't Close " & f.Name
    Unload f
    Set f = Nothing
    End If
    
    Next f
    
    Call Disconnect_Database

Set waFileSub = Nothing
Set waMasterSub = Nothing
Set waOperationSub = Nothing
Set waPOSub = Nothing
Set waInventorySub = Nothing
Set waACCOUNTSub = Nothing
Set waInquirySub = Nothing
Set waReportSub = Nothing
Set waUtilitySub = Nothing
Set waAccRptSub = Nothing
Set waListSub = Nothing
Set waScrItm = Nothing
Set waScrToolTip = Nothing
Set nbMain = Nothing

End Sub

Private Sub lvwDB_ColumnClick(ByVal columnheader As columnheader)
    If lvwDB.SortOrder = lvwAscending Then
    lvwDB.SortOrder = lvwDescending
    Else
    lvwDB.SortOrder = lvwAscending
    End If
    lvwDB.SortKey = columnheader.Index - 1
    ' Set Sorted to True to sort the list.
    lvwDB.Sorted = True
    
End Sub

Private Sub lvwDB_DblClick()
  
Dim wsFName As String

If lvwDB.ListItems.Count = 0 Then
Exit Sub
End If

If lvwDB.SelectedItem Is Nothing Then
Exit Sub
End If

Me.MousePointer = vbHourglass

 wsFName = lvwDB.SelectedItem.Text
    
 Call Call_Pgm(waFileSub, 0, UCase(wsFName))
    
    
   
Me.MousePointer = vbNormal
   
End Sub




Private Sub lvwDB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        lvwDB_DblClick
        
        
    End If
End Sub

Private Sub lvwDB_LostFocus()
For i = 1 To lvwDB.ListItems.Count
      lvwDB.ListItems.Item(i).Selected = False
Next i
      
End Sub


Private Sub mnuACCRPTSub_Click(Index As Integer)
    Call Call_Pgm(waAccRptSub, Index)
End Sub

Private Sub mnuListSub_Click(Index As Integer)
    Call Call_Pgm(waListSub, Index)
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Call Call_Pgm(waFileSub, Index)
End Sub

Private Sub mnuMasterSub_Click(Index As Integer)
    Call Call_Pgm(waMasterSub, Index)
End Sub

Private Sub mnuOperationSub_Click(Index As Integer)
    Call Call_Pgm(waOperationSub, Index)
End Sub

Private Sub mnuInventorySub_Click(Index As Integer)
    Call Call_Pgm(waInventorySub, Index)
End Sub
Private Sub mnuinquirySub_Click(Index As Integer)
    Call Call_Pgm(waInquirySub, Index)
End Sub



Private Sub mnuPOSub_Click(Index As Integer)
    Call Call_Pgm(waPOSub, Index)
End Sub

Private Sub mnureportSub_Click(Index As Integer)
    Call Call_Pgm(waReportSub, Index)
End Sub

Private Sub mnuACCOUNTSub_Click(Index As Integer)
    Call Call_Pgm(waACCOUNTSub, Index)
End Sub

Private Sub mnuutilitySub_Click(Index As Integer)
    Call Call_Pgm(waUtilitySub, Index)
End Sub


Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
        Case tcPrev
            
        If tbrMain.Buttons(tcPrev).Enabled = True Then
            
            ChgPrevPage
            
        End If
        
        Case tcNext
        
         If tbrMain.Buttons(tcNext).Enabled = True Then
              
            ChgNextPage
            
         End If
         
        Case tcOK
            
            Call Call_Pgm(waFileSub, 0, UCase(cboCommand.Text))
            
        Case tcCancel
            
            Call cmdCancel
            
        Case tcExit
            
           
            Unload Me
            
    End Select
End Sub

Private Sub cmdCancel()

            cboCommand.Clear
            cboCommand.Text = ""
            giFormIndex = 0
                With tbrMain
                .Buttons(tcPrev).Enabled = False
                .Buttons(tcNext).Enabled = False
             End With

End Sub
    

Private Sub tvwDB_Collapse(ByVal Node As Node)
    If Node.Tag = "Root" Or Node.Index = 1 _
    Then Node.Image = "closed"
End Sub

Private Sub tvwDB_Expand(ByVal Node As Node)
    If Node.Tag = "Root" Or Node.Index = 1 Then
        Node.Image = "open"
     '   Node.Sorted = True
    End If
    
      
   
End Sub




Private Sub GetItem(ITMID, ByVal inArray As XArrayDB, ByVal sImage As String)

Dim wiCtr As Integer
    ' Show Progress bar
     ' Clear the old titles
    lvwDB.ListItems.Clear
                
            
            
        For wiCtr = 0 To inArray.UpperBound(1)
                
        If inArray(wiCtr, 2) = "Y" And inArray(wiCtr, 1) <> "-" Then
        
            Set mItem = lvwDB.ListItems.Add _
            (, , inArray(wiCtr, 0), sImage, sImage)
            
            mItem.SubItems(1) = SkipA(inArray(wiCtr, 1))
        End If
            
        Next wiCtr
    
    
    sCurrentIndex = ITMID

End Sub


Private Sub tvwDB_NodeClick(ByVal Node As Node)
    ' Check the Tag for "Publisher" and EventFlag
    ' variable to see if the ColumnHeaders
    ' have already been created. If not, then
    ' invoke the MakeColumns procedure.
  
    ' If the Tag is "Publisher" and the mItemCurrentIndex
    ' index isn't the same as the Node.key, then
    ' incoke the GetItem procedure.
    If Node.Tag = "Item" And sCurrentIndex <> Node.Key Then
    
    Select Case Node.Key
    Case "FILE"
    GetItem Node.Key, waFileSub, Node.Key
    Case "MASTER"
    GetItem Node.Key, waMasterSub, Node.Key
    Case "OPERATION"
    GetItem Node.Key, waOperationSub, Node.Key
    Case "PO"
    GetItem Node.Key, waPOSub, Node.Key
    Case "INVENTORY"
    GetItem Node.Key, waInventorySub, Node.Key
    Case "ACCOUNT"
    GetItem Node.Key, waACCOUNTSub, Node.Key
    Case "INQUIRY"
    GetItem Node.Key, waInquirySub, Node.Key
    Case "REPORT"
    GetItem Node.Key, waReportSub, Node.Key
    Case "UTILITY"
    GetItem Node.Key, waUtilitySub, Node.Key
    Case "ACCRPT"
    GetItem Node.Key, waAccRptSub, Node.Key
    Case "LIST"
    GetItem Node.Key, waListSub, Node.Key
    
    End Select
    End If
    
End Sub

Private Sub FindInList()   ' FindItem method.

 Dim intSelectedOption As Integer
 Dim strFindMe As String
      
      strFindMe = txtOutput.Text
      
      intSelectedOption = lvwText
     
      'lvwSubitem
      'lvwText
   
   Set mItem = lvwDB. _
   FindItem(strFindMe, intSelectedOption, , lvwPartial)
   If mItem Is Nothing Then  ' If no match, inform user and exit.
    '  MsgBox "No match found"
      Exit Sub
    Else
       mItem.EnsureVisible ' Scroll ListView to show found ListItem.
       mItem.Selected = True   ' Select the ListItem.
       lvwDB.SetFocus
   End If
   
   End Sub



Private Sub LoadForm(f As Form)
   f.WindowState = 0
   f.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight
   f.Show
   f.ZOrder 0
   
End Sub


Private Sub Call_Pgm(ByVal inArray As XArrayDB, inPgmIdx As Integer, Optional inPgmName, Optional inNotAdd)

    Dim newForm As Form
    Dim wsFName As String
    
    On Error GoTo Err_Handler
    
    If IsMissing(inPgmName) Then
        wsFName = inArray(inPgmIdx, 0)
    Else
        wsFName = inPgmName
    End If
    
    
   If Chk_PgmExist(wsFName) = False Then
            gsMsg = "畫面不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCommand.SetFocus
            Exit Sub
   End If
        
        
   If Chk_UserRight(gsUserID, UCase(wsFName)) = False Then
            gsMsg = "使用者權限不足!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCommand.SetFocus
            Exit Sub
   End If
    
Me.MousePointer = vbHourglass
    

    Select Case wsFName
        Case "EXIT"
           Unload Me
        
        Case "CMP001"
            Set newForm = New frmCMP001
            newForm.Show vbModal
                    
        Case "SYS001"
            Set newForm = New frmSYS001
            newForm.Show vbModal
        
        Case "SYS002"
            Set newForm = New frmSYS002
            newForm.Show
                   
        Case "WS001"
            Set newForm = New frmWS001
            newForm.Show vbModal
        
        Case "VOU001"
            Set newForm = New frmVOU001
            newForm.Show
        
        
        Case "OPN001"
            Set newForm = New frmOPN001
            newForm.Show vbModal
    
        Case "OPN002"
            Set newForm = New frmOPN002
            newForm.Show vbModal
            
        Case "UC001"
            Set newForm = New frmUC001
            newForm.Show vbModal
            
            
 ''''
        Case "C001"
            Set newForm = New frmC001
            newForm.Show
    
       Case "V001"
            Set newForm = New frmV001
            newForm.Show
 
       Case "SLM001"
            Set newForm = New frmSLM001
            newForm.Show
 
       Case "STF001"
            Set newForm = New frmSTF001
            newForm.Show
  
        
        Case "ITM001"
            Set newForm = New frmITM001
            newForm.Show
 

        
        Case "PYT001"
            Set newForm = New frmPYT001
            newForm.Show
                     
        Case "PR001"
            Set newForm = New frmPR001
            newForm.Show
        
        Case "EXC001"
            Set newForm = New frmEXC001
            newForm.Show
        
        Case "UOM001"
            Set newForm = New frmUOM001
            newForm.Show
            
        Case "IP001"
            Set newForm = New frmIP001
            newForm.Show
            
        Case "SHP001"
            Set newForm = New frmSHP001
            newForm.Show
            
        Case "RMK001"
            Set newForm = New frmRmk001
            newForm.Show
            
        Case "WH001"
            Set newForm = New frmWH001
            newForm.Show
            
        Case "IT001"
            Set newForm = New frmIT001
            newForm.Show
            
        Case "PT001"
            Set newForm = New frmPT001
            newForm.Show
            
        Case "RGN001"
            Set newForm = New frmRGN001
            newForm.Show
            
 
        Case "COA001"
            Set newForm = New frmCOA001
            newForm.Show
            
        Case "ML001"
            Set newForm = New frmML001
            newForm.Show
 
        Case "M001"
            Set newForm = New frmM001
            newForm.Show
 
         Case "N001"
            Set newForm = New frmN001
            newForm.Show
 
         Case "TERR001"
            Set newForm = New frmTerr001
            newForm.Show
 
 
         Case "SHP001"
            Set newForm = New frmSHP001
            newForm.Show
            
         Case "AT001"
         Set newForm = New frmAT001
            newForm.Show
            
        Case "CAT001"
         Set newForm = New frmCAT001
            newForm.Show
 
 ''''
        Case "SN001"
         Set newForm = New frmSN001
            newForm.Show
            
        Case "SO001"
         Set newForm = New frmSO001
            newForm.Show
            
            
        Case "VQ001"
         Set newForm = New frmVQ001
            newForm.Show
            
            
   ''     Case "SPL001"
    ''     Set newForm = New frmSPL001
     ''       newForm.Show
            
        Case "SDN001"
         Set newForm = New frmSDN001
            newForm.Show
        
        Case "INV001"
         Set newForm = New frmINV001
            newForm.Show
        
        
        Case "APR001"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR001"
            .TrnCd = "SN"
            .Show
         End With
            
            
        Case "APR002"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR002"
            .TrnCd = "SO"
            .Show
         End With
                        
       Case "APR003"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR003"
            .TrnCd = "SP"
            .Show
         End With
                         
       Case "APR004"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR004"
            .TrnCd = "SD"
            .Show
         End With
         
       Case "APR005"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR005"
            .TrnCd = "IV"
            .Show
         End With
         
         
       Case "APR006"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR006"
            .TrnCd = "SW"
            .Show
         End With
         
       Case "APR007"
         Set newForm = New frmAPR001
         With newForm
            .FormID = "APR007"
            .TrnCd = "VQ"
            .Show
         End With
         
        Case "APW001"
         Set newForm = New frmAPW001
         With newForm
            .FormID = "APW001"
             .Show
         End With
                  
                         
 '''' Purchase Order
 
        Case "MRP001"
         Set newForm = New frmMRP001
         newForm.Show
       
        Case "MRP002"
         Set newForm = New frmMRP002
         With newForm
            .FormID = "MRP002"
             .Show
         End With
        
        Case "PO001"
         Set newForm = New frmPO001
         With newForm
            .FormID = "PO001"
             .Show
         End With
            
        Case "PN001"
         Set newForm = New frmPO001
         With newForm
            .FormID = "PN001"
             .Show
         End With
            
       Case "PGV001"
         Set newForm = New frmPGV001
            newForm.Show
            
       Case "GRV001"
         Set newForm = New frmGRV001
            newForm.Show
            
        
       Case "PGR001"
         Set newForm = New frmPGR001
            newForm.Show
        
        Case "APV001"
         Set newForm = New frmAPV001
         With newForm
            .FormID = "APV001"
            .TrnCd = "PO"
            .Show
         End With
            
        Case "APP001"
         Set newForm = New frmAPP001
         With newForm
            .FormID = "APP001"
             .Show
         End With
            
            
        Case "APV002"
         Set newForm = New frmAPV001
         With newForm
            .FormID = "APV002"
            .TrnCd = "PV"
            .Show
         End With
                        
       Case "APV003"
         Set newForm = New frmAPV001
         With newForm
            .FormID = "APV003"
            .TrnCd = "PR"
            .Show
         End With
         

       Case "APV004"
         Set newForm = New frmAPV001
         With newForm
            .FormID = "APV004"
            .TrnCd = "GR"
            .Show
         End With

 ''''''''''''Inventory
 
         
       Case "TRF001"
         Set newForm = New frmTRF001
         newForm.Show
            
       Case "ADJ001"
         Set newForm = New frmADJ001
         newForm.Show
         
       Case "SAM001"
         Set newForm = New frmSAM001
         newForm.Show
            
       Case "DAM001"
         Set newForm = New frmDAM001
         newForm.Show
            
       Case "SKT001"
         Set newForm = New frmSKT001
         newForm.Show
            
       Case "SCT001"
         Set newForm = New frmSCT001
         newForm.Show
            
            
        Case "APS001"
         Set newForm = New frmAPS001
         With newForm
            .FormID = "APS001"
            .Show
         End With
            
 ''''''''''''
 
        Case "USR001"
            Set newForm = New frmUSR001
            newForm.Show
        
        
        Case "CHGPWD"
            Set newForm = New frmCHGPWD
            newForm.Show vbModal
        
        
        Case "USRRHT"
            Set newForm = New frmUSRRHT
            newForm.Show vbModal
        
        Case "PURGE"
            Set newForm = New frmPURGE
            newForm.Show vbModal

        Case "HHIM001"
            Set newForm = New frmHHIM001
            newForm.Show vbModal


        '-----------AR
        
        Case "AR001"
            Set newForm = New frmAR001
            With newForm
                .FormID = "AR001"
                .TrnCd = "62"
                .Show
            End With
            
        Case "ARDN001"
            Set newForm = New frmAR001
            With newForm
                .FormID = "ARDN001"
                .TrnCd = "61"
                .Show
            End With
            
        Case "ARCN001"
            Set newForm = New frmAR001
            With newForm
                .FormID = "ARCN001"
                .TrnCd = "60"
                .Show
            End With
            
        Case "AR002"
            Set newForm = New frmAR002
            newForm.Show
        
        Case "AR003"
            Set newForm = New frmAR003
            newForm.Show
            
        Case "AR003"
            Set newForm = New frmAR003
            newForm.Show
            
        Case "AR100"
            Set newForm = New frmAR100
            newForm.Show
            
        Case "AR101"
            Set newForm = New frmAR101
            newForm.Show
            
        Case "ARPE000"
            Set newForm = New frmARPE000
            newForm.Show
            
         '-----------AP
        
        Case "AP001"
            Set newForm = New frmAP001
            With newForm
                .FormID = "AP001"
                .TrnCd = "20"
                .Show
            End With
            
        Case "APCN001"
            Set newForm = New frmAP001
            With newForm
                .FormID = "APCN001"
                .TrnCd = "21"
                .Show
            End With
        
        Case "AP002"
            Set newForm = New frmAP002
            newForm.Show
        
        Case "AP003"
            Set newForm = New frmAP003
            newForm.Show
            
        Case "AP003"
            Set newForm = New frmAP003
            newForm.Show
            
        Case "AP100"
            Set newForm = New frmAP100
            newForm.Show
            
        Case "AP101"
            Set newForm = New frmAP101
            newForm.Show
        
        Case "SIGN002"
            Set newForm = New frmSIGN002
            newForm.Show
        
         '-----------GL
        
        Case "GL001"
            Set newForm = New frmGL001
            newForm.Show
        
        Case "GL002"
            Set newForm = New frmGL002
            newForm.Show
        
        '----------Acc Prt
        
        Case "VOU002"
            Set newForm = New frmVOU002
            newForm.Show vbModal
        
        Case "COA002"
            Set newForm = New frmCOA002
            newForm.Show vbModal
        
        
        '-----------
        
        Case "ARL001"
            Set newForm = New frmARL001
            newForm.Show vbModal
        
        Case "ARL002"
            Set newForm = New frmARL002
            newForm.Show vbModal
        
        Case "ARL003"
            Set newForm = New frmARL003
            newForm.Show vbModal
        
        Case "ARL004"
            Set newForm = New frmARL004
            newForm.Show vbModal
        
        Case "ARL005"
            Set newForm = New frmARL005
            newForm.Show vbModal
        
        Case "ARL006"
            Set newForm = New frmARL006
            newForm.Show vbModal
        
        Case "ARL007"
            Set newForm = New frmARL007
            newForm.Show vbModal
        
        Case "ARL008"
            Set newForm = New frmARL008
            newForm.Show vbModal
        
        Case "ARL009"
            Set newForm = New frmARL009
            newForm.Show vbModal
        
        Case "ARL010"
            Set newForm = New frmARL010
            newForm.Show vbModal
        
        Case "ARL011"
            Set newForm = New frmARL011
            newForm.Show vbModal
        
        Case "ARL012"
            Set newForm = New frmARL012
            newForm.Show vbModal
       '-----------
        
        Case "APL001"
            Set newForm = New frmAPL001
            newForm.Show vbModal
        
        Case "APL002"
            Set newForm = New frmAPL002
            newForm.Show vbModal
        
        Case "APL003"
            Set newForm = New frmAPL003
            newForm.Show vbModal
        
        Case "APL004"
            Set newForm = New frmAPL004
            newForm.Show vbModal
        
        Case "APL005"
            Set newForm = New frmAPL005
            newForm.Show vbModal
        
        Case "APL006"
            Set newForm = New frmAPL006
            newForm.Show vbModal
        
        Case "APL007"
            Set newForm = New frmAPL007
            newForm.Show vbModal
        
        Case "APL008"
            Set newForm = New frmAPL008
            newForm.Show vbModal
        
        Case "APL009"
            Set newForm = New frmAPL009
            newForm.Show vbModal
         
        Case "APL010"
            Set newForm = New frmAPL010
            newForm.Show vbModal
            
       '-----------------------
        Case "GLP001"
            Set newForm = New frmGLP001
            newForm.Show vbModal
            
        Case "GLP002"
            Set newForm = New frmGLP002
            newForm.Show vbModal
        
        Case "GLP003"
            Set newForm = New frmGLP003
            newForm.Show vbModal
        
        Case "GLP004"
            Set newForm = New frmGLP004
            newForm.Show vbModal
        
        Case "GLP005"
            Set newForm = New frmGLP005
            newForm.Show vbModal
                
        Case "GLP006"
            Set newForm = New frmGLP006
            newForm.Show vbModal
''' Job Cost
        
        
 '       Case "CST001"
 '           Set newForm = New frmCST001
 '           newForm.Show
            
 '       Case "CT001"
 '           Set newForm = New frmCT001
 '           newForm.Show
            
        ''' Reporting
        
         Case "SN002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SN002"
            .TrnCd = "SN"
            .Show
         End With
         
         
         Case "SN002D"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SN002D"
            .TrnCd = "SN"
            .Show
         End With
         
         Case "SO002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SO002"
            .TrnCd = "SO"
            .Show
         End With
            
         
         Case "SO002D"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SO002D"
            .TrnCd = "SO"
            .Show
         End With
         
         Case "SPL002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SPL002"
            .TrnCd = "SP"
            .Show
         End With
                    
            
         Case "SDN002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "SDN002"
            .TrnCd = "SD"
            .Show
         End With
         
         
         Case "INV002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "INV002"
            .TrnCd = "IV"
            .Show
         End With
        
        Case "PO002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "PO002"
            .TrnCd = "PO"
            .Show
         End With
                
        Case "PGV002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "PGV002"
            .TrnCd = "PV"
            .Show
         End With
         
         Case "GRV002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "GRV002"
            .TrnCd = "GR"
            .Show
         End With
         
        
        Case "PGR002"
         Set newForm = New frmSN002
         With newForm
            .FormID = "PGR002"
            .TrnCd = "PR"
            .Show
         End With
        
 ''''''''''''''''''
        Case "JOB001"
         Set newForm = New frmJOB002
         With newForm
            .FormID = "JOB001"
            .Show
         End With
        
        
        Case "JOB002"
         Set newForm = New frmJOB002
         With newForm
            .FormID = "JOB002"
            .Show
         End With
         
        
        Case "JOB003"
         Set newForm = New frmJOB002
         With newForm
            .FormID = "JOB003"
            .Show
         End With
         
       
        Case "JOB004"
         Set newForm = New frmJOB002
         With newForm
            .FormID = "JOB004"
            .Show
         End With
       
       ' Case "SOP000B"
       '     Set newForm = New frmSOP000B
       '     newForm.Show
            
        Case "SIGN001"
            Set newForm = New frmSIGN001
            newForm.Show
        
        Case "ITM002"
            Set newForm = New frmITM002
            newForm.Show
            
 ''''' Sales Report
            
        Case "SOP000"
            Set newForm = New frmSOP000
            newForm.Show
            
        Case "SOP003"
            Set newForm = New frmSOP003
            newForm.Show
            
        Case "SOP004"
            Set newForm = New frmSOP004
            newForm.Show
 
        Case "SOP006"
            Set newForm = New frmSOP006
            newForm.Show
 
        Case "SOP008"
            Set newForm = New frmSOP008
            newForm.Show
 
        Case "SOP010"
            Set newForm = New frmSOP010
            newForm.Show
 
        Case "SOP020"
            Set newForm = New frmSOP020
            newForm.Show
 
        Case "SOP030"
            Set newForm = New frmSOP030
            newForm.Show
            
        Case "ICP001"
            Set newForm = New frmICP001
            newForm.Show
            
        Case "ICP002"
            Set newForm = New frmICP002
            newForm.Show
            
        Case "ICP003"
            Set newForm = New frmICP003
            newForm.Show
            
        Case "ICP004"
            Set newForm = New frmICP004
            newForm.Show
                        
        Case "ICP005"
            Set newForm = New frmICP005
            newForm.Show
            
        Case "ICP006"
            Set newForm = New frmICP006
            newForm.Show
 
 '''' Inquiry
 
        Case "INQ001"
            Set newForm = New frmINQ001
            With newForm
                .FormID = "INQ001"
                .TrnCd = "SO"
                .Show
            End With
 
         Case "INQ002"
            Set newForm = New frmINQ001
            With newForm
                .FormID = "INQ002"
                .TrnCd = "IV"
                .Show
            End With
 
           
        Case "INQ003"
            Set newForm = New frmINQ003
            newForm.Show
         
         Case "INQ004"
            Set newForm = New frmINQ001
            With newForm
                .FormID = "INQ004"
                .TrnCd = "PO"
                .Show
            End With
            
         Case "INQ005"
            Set newForm = New frmINQ001
            With newForm
                .FormID = "INQ005"
                .TrnCd = "PV"
                .Show
            End With
            
            
        Case "INQ006"
            Set newForm = New frmINQ006
            newForm.Show
            
        Case "INQ007"
            Set newForm = New frmINQ007
            newForm.Show
            
        Case "INQ008"
            Set newForm = New frmINQ008
            newForm.Show
            
        Case "INQ009"
            Set newForm = New frmINQ009
            newForm.Show
            
        Case "INQ010"
            Set newForm = New frmINQ010
            newForm.Show
            
         Case "INQ011"
            Set newForm = New frmINQ001
            With newForm
                .FormID = "INQ011"
                .TrnCd = "SN"
                .Show
            End With
            
        Case "INQ012"
            Set newForm = New frmINQ012
            newForm.Show
            
        Case "INQ013"
            Set newForm = New frmINQ013
            newForm.Show
            
        Case "STKCNT"
            Set newForm = New frmSTKCNT
            newForm.Show
            
 '''' Master Listing
        Case "AT002"
            Set newForm = New frmAT002
            newForm.Show
        
        Case "C002"
            Set newForm = New frmC002
            newForm.Show
            
        Case "COA002"
            Set newForm = New frmCOA002
            newForm.Show
        
        Case "EXC002"
            Set newForm = New frmEXC002
            newForm.Show
        
        Case "IP0022"
            Set newForm = New frmIP0022
            newForm.Show
            
        Case "IT002"
            Set newForm = New frmIT002
            newForm.Show
            
        Case "ML002"
            Set newForm = New frmML002
            newForm.Show
            
        Case "PR002"
            Set newForm = New frmPR002
            newForm.Show
            
        Case "PT002"
            Set newForm = New frmPT002
            newForm.Show
            
        Case "PYT002"
            Set newForm = New frmPYT002
            newForm.Show
            
        Case "RMK002"
            Set newForm = New frmRMK002
            newForm.Show
            
        Case "SHP002"
            Set newForm = New frmSHP002
            newForm.Show
            
        Case "SLM002"
            Set newForm = New frmSLM002
            newForm.Show
            
        Case "STF002"
            Set newForm = New frmSTF002
            newForm.Show
            
        Case "UOM002"
            Set newForm = New frmUOM002
            newForm.Show
            
        Case "USR002"
            Set newForm = New frmUSR002
            newForm.Show
            
        Case "V002"
            Set newForm = New frmV002
            newForm.Show
            
        Case "WH002"
            Set newForm = New frmWH002
            newForm.Show
        
        Case Else
            Me.MousePointer = vbNormal
            Exit Sub
            
    End Select
    
    
   If IsMissing(inNotAdd) Then
    cboCommand.AddItem wsFName
    tbrMain.Buttons(tcPrev).Enabled = True
    giCurrIndex = giCurrIndex + 1
   End If
    
   Me.MousePointer = vbNormal
          
Exit Sub

Err_Handler:

   Me.MousePointer = vbNormal

    
End Sub

