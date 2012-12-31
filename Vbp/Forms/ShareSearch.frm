VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShareSearch 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   6735
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "ShareSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   6731.463
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   7703.93
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSearchField 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin MSComctlLib.ListView lsvContent 
      Height          =   5055
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   8916
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraFind 
      Caption         =   "搜尋"
      Height          =   810
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   7260
      Begin VB.TextBox txtOutput 
         Height          =   270
         Left            =   3240
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   7200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":1E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":21B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":2606
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":2A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":2D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":308C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":34DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":3DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShareSearch.frx":40E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDspCrt 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   7305
   End
End
Attribute VB_Name = "frmShareSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim wrsTbl As rdoResultset
Dim clmX As columnheader
Dim itmX As ListItem
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB


'Criteria variables
Public sBindSQL As String                   '-- SQL to bind to the ADO control.
Public sBindWhereSQL As String
Public sBindOrderSQL As String
Private wsWhereSql As String

'List header array
Public vHeadDataAry As Variant              '-- Two dimension arary. 1-Field description; 2-Field name.
'Filter field combo array
Public vFilterAry As Variant                '-- Two dimension arary. 1-Field description; 2-Field name.
'Search field combo array
Public vSearchAry As Variant                '-- Two dimension arary. 1-Field description; 2-Field name.

Private wsFormCaption As String
Private wsFormID As String
Private wiExit As Boolean

Private Const tcGo = "Go"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"


Private Sub cboSearchField_GotFocus()
    FocusMe cboSearchField
End Sub

Private Sub cboSearchField_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtOutput.SetFocus
    End If
End Sub

Private Sub cboSearchField_LostFocus()
FocusMe cboSearchField, True
End Sub

Private Sub Form_Deactivate()

    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            lsvContent_DblClick
            
        Case vbKeyF3
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
        
        Case vbKeyF5
            Edt_Lst
        
        Case vbKeyEscape
            KeyCode = vbDefault
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    
    
  MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Scr
    IniCombo
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCancel()
    
    
  MousePointer = vbHourglass
  
    Ini_Scr
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub IniCombo()
    Dim wiCounter As Integer
    
    With cboSearchField
        For wiCounter = 1 To UBound(vFilterAry)
            .AddItem vFilterAry(wiCounter, 1)
        Next
    End With
    
    cboSearchField.ListIndex = 0
    

End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
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
    wsWhereSql = ""
    wiExit = False
        
    Call Ini_LstView
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If wiExit = False Then
       Cancel = True
       wiExit = True
       Me.Hide
       Exit Sub
    End If

    Set clmX = Nothing
    Set itmX = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set frmShareSearch = Nothing
    
    
End Sub

Private Sub lsvContent_DblClick()
    If Not lsvContent.SelectedItem Is Nothing Then
        Me.Tag = lsvContent.SelectedItem
    Else
        gsMsg = "請先選取一項!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Sub
    End If
    'Me.Hide
    Unload Me
End Sub

Private Sub lsvContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Call lsvContent_DblClick
    End If
End Sub

Private Sub Edt_Lst()
    Dim Criteria As String
    Dim inpData As Variant
    Dim wsRows As String
    Dim rsSearch As New ADODB.Recordset
    Dim iCounter As Integer
        
    On Error GoTo Edt_Lst_Err
    
    lsvContent.Enabled = True
    lsvContent.ListItems.Clear
    lsvContent.Refresh
    
    MousePointer = vbHourglass
    
    'Criteria = Criteria & " CusName LIKE '%" & Set_Quote(txtCustomerName.Text) & "%'"
    
    Criteria = sBindSQL & sBindWhereSQL & wsWhereSql & sBindOrderSQL
    
    rsSearch.Open Criteria, cnCon, adOpenStatic, adLockOptimistic
    
    If rsSearch.RecordCount = 0 Then
        lsvContent.Enabled = False
        rsSearch.Close
        Set rsSearch = Nothing
        MousePointer = vbDefault
        Exit Sub
    End If
    
    rsSearch.MoveFirst
    
    Do Until rsSearch.EOF
        'inpData = rsSearch("CusCode").Value
        inpData = rsSearch.Fields.Item(0).Value
        Set itmX = lsvContent.ListItems.Add(, , inpData)
        
        'inpData = rsSearch("CusName").Value
        For iCounter = 1 To rsSearch.Fields.Count - 1
            inpData = rsSearch.Fields.Item(iCounter).Value
            itmX.SubItems(iCounter) = IIf(IsNull(inpData), "", inpData)
        Next
    
        rsSearch.MoveNext
    Loop
        
    lsvContent.SetFocus
    
    MousePointer = vbDefault

    Exit Sub
    
Edt_Lst_Err:
    'Dsp_Err "", Err.Description, "E", Me.Caption
    gsMsg = "Err in Edt_Lst_Err!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    MousePointer = vbDefault
End Sub

Private Sub Ini_LstView()
    Dim iCounter As Integer
    
    lsvContent.ListItems.Clear
    lsvContent.ColumnHeaders.Clear
    
    With lsvContent
        Set clmX = .ColumnHeaders. _
            Add(, , vHeadDataAry(1, 1), vHeadDataAry(1, 3), lvwColumnLeft)
        clmX.Alignment = lvwColumnLeft
        
        For iCounter = 2 To UBound(vHeadDataAry)
            Set clmX = .ColumnHeaders. _
                Add(, , vHeadDataAry(iCounter, 1), vHeadDataAry(iCounter, 3), lvwColumnCenter)
            clmX.Alignment = lvwColumnLeft
        Next
        
        .BorderStyle = ccFixedSingle    ' Set BorderStyle property.
        .View = lvwReport               ' Set View property to Report.
        .Font.Name = "MS Sans Serif"
        .Font.Bold = False
        .Font.Size = 8
        .ForeColor = &HC00000
        .DragMode = 0
        .Sorted = False
    End With
    lsvContent.ListItems.Clear
    lsvContent.Enabled = False
End Sub

Private Sub lsvContent_ColumnClick(ByVal columnheader As columnheader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1

    lsvContent.SortKey = columnheader.Index - 1
    lsvContent.SortOrder = 1 - lsvContent.SortOrder
    ' Set Sorted to True to sort the list.
    lsvContent.Sorted = True
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "ShareSrch"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F2)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    
    fraFind.Caption = Get_Caption(waScrItm, "FRAFIND")
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Go"
            lsvContent_DblClick
            
        Case "Cancel"
        
           Call cmdCancel
            
        Case "Exit"
            Unload Me
            
        Case "Refresh"
            Edt_Lst
            
    End Select
End Sub

Private Sub txtOutput_GotFocus()
    FocusMe txtOutput
End Sub



Private Sub setCriteria1()
    
    If InStr(1, UCase(sBindWhereSQL), vFilterAry(cboSearchField.ListIndex + 1, 2), vbTextCompare) <> 0 Then
        Exit Sub
    End If
    
    If InStr(1, UCase(wsWhereSql), vFilterAry(cboSearchField.ListIndex + 1, 2), vbTextCompare) <> 0 Then
        wsWhereSql = ""
        lblDspCrt.Caption = ""
    End If
    
    If Trim(txtOutput) = "" Then
        Exit Sub
    End If

    If InStr(1, UCase(sBindWhereSQL), "WHERE", vbTextCompare) <> 0 Then
        wsWhereSql = wsWhereSql & " AND " + vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '%" & Set_Quote(txtOutput.Text) & "%'"
    Else
        If InStr(1, UCase(wsWhereSql), "WHERE", vbTextCompare) <> 0 Then
            wsWhereSql = wsWhereSql & " AND " + vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '%" & Set_Quote(txtOutput.Text) & "%'"
        Else
            wsWhereSql = " WHERE " + vFilterAry(cboSearchField.ListIndex + 1, 2) & " LIKE '%" & Set_Quote(txtOutput.Text) & "%'"
        End If
    End If
    
    lblDspCrt.Caption = lblDspCrt.Caption & " " & vFilterAry(cboSearchField.ListIndex + 1, 1) & " = " & Set_Quote(txtOutput.Text) & ": "
    
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cboSearchField.SetFocus
        setCriteria1
        Edt_Lst
       
    End If
End Sub

Private Sub txtOutput_LostFocus()
    FocusMe txtOutput, True
End Sub
