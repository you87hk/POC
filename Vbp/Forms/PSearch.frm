VERSION 5.00
Object = "{9FE255D1-F32E-11D0-9E15-444553540000}#1.0#0"; "MLISTX.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPSearch 
   BorderStyle     =   1  '單線固定
   Caption         =   "快速搜尋"
   ClientHeight    =   7035
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   6510
   Icon            =   "PSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6510
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6120
      Picture         =   "PSearch.frx":0442
      ScaleHeight     =   330
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6240
      Begin VB.CommandButton cmdQuickAdd 
         Height          =   255
         Left            =   5880
         Picture         =   "PSearch.frx":05CC
         Style           =   1  '圖片外觀
         TabIndex        =   6
         ToolTipText     =   "Quick Add"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtOutput 
         Height          =   270
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label frmSearch1 
         Caption         =   "搜尋 :"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1920
      End
   End
   Begin MabryCtl.MList MList1 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6255
      _ExtentX        =   5080
      _ExtentY        =   5080
      BorderEffect    =   2
      CaptionEffect   =   3
      ColDelimiter    =   "|"
      HeadingsEffect  =   4
      Object.TabStop         =   -1  'True
      ColRowOrder     =   -1  'True
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Columns {23BAA6DE-05A6-11D1-9E15-0020AFD6A9D5} 
         ColumnCount     =   5
         BeginProperty Column0 {23BAA6E0-05A6-11D1-9E15-0020AFD6A9D5} 
            Object.Width           =   0
            MinWidth        =   0
            MaxWidth        =   -1
            UserResizeEnabled=   -1
            Heading         =   "Col1"
            Object.Visible         =   -1
            ColumnAlignment =   0
            HeadingAlignment=   0
         EndProperty
         BeginProperty Column1 {23BAA6E0-05A6-11D1-9E15-0020AFD6A9D5} 
            Object.Width           =   0
            MinWidth        =   0
            MaxWidth        =   -1
            UserResizeEnabled=   -1
            Heading         =   "Col 2"
            Object.Visible         =   -1
            ColumnAlignment =   0
            HeadingAlignment=   0
         EndProperty
         BeginProperty Column2 {23BAA6E0-05A6-11D1-9E15-0020AFD6A9D5} 
            Object.Width           =   0
            MinWidth        =   0
            MaxWidth        =   -1
            UserResizeEnabled=   -1
            Heading         =   "Col 3"
            Object.Visible         =   -1
            ColumnAlignment =   0
            HeadingAlignment=   0
         EndProperty
         BeginProperty Column3 {23BAA6E0-05A6-11D1-9E15-0020AFD6A9D5} 
            Object.Width           =   0
            MinWidth        =   0
            MaxWidth        =   -1
            UserResizeEnabled=   -1
            Heading         =   "Col 4"
            Object.Visible         =   -1
            ColumnAlignment =   0
            HeadingAlignment=   0
         EndProperty
         BeginProperty Column4 {23BAA6E0-05A6-11D1-9E15-0020AFD6A9D5} 
            Object.Width           =   0
            MinWidth        =   0
            MaxWidth        =   -1
            UserResizeEnabled=   -1
            Heading         =   "Col 5"
            Object.Visible         =   -1
            ColumnAlignment =   0
            HeadingAlignment=   0
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
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
            Picture         =   "PSearch.frx":073E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":1018
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":1334
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":1C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":2060
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":27CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":2C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":3070
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":338A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PSearch.frx":36A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
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
End
Attribute VB_Name = "frmPSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sBindSQL As String                   '-- SQL to bind to the ADO control.

'Public sCountSQL As String                  '-- Display total no. of records SQL.
Public vHeadDataAry As Variant              '-- Two dimension arary. 1-Field description; 2-Field name.
Public oTextBox As TextBox               '-- Calling form ADO control reference to locate record purpose.
Public lKeyPos As Long                   '-- For Recordset.Find purpose.
Public lSearchPos As Long                   '-- Check the type of sKeyName (String or Numeric).
Public sInputType

Private Sub cmdQuickAdd_Click()
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sUpdSQL As String                   '-- For update PO Items.
    Dim sWarehouse As String
    Dim sResult As String
    
    If txtOutput.Text = "" Then
        sMsg = "沒有輸入須要之資料!"
        MsgBox sMsg, vbInformation + vbOKOnly, gsTitle
        Exit Sub
    End If
    
    Set rsRcd = cnCon.Execute("SELECT * from tblInput WHERE InputName='" & txtOutput & "'")
    If Not rsRcd.EOF Then
        sMsg = "資料找不到, 請重新整理再選取想要之資料!"
        MsgBox sMsg, vbInformation + vbOKOnly, gsTitle
        rsRcd.Close
        Exit Sub
    End If
    rsRcd.Close
    
    sSQL = "INSERT INTO tblInput(InputName, InputType, "
    sSQL = sSQL & "UpdUser, UpdTime) VALUES('" & txtOutput & "','" & sInputType
    sSQL = sSQL & "', '" & gsUserID & "', '" & gsSystemDate & Time & "')"
    
    cnCon.BeginTrans
    '-- Add record to warehouse table.
    cnCon.Execute sSQL
    
    '-- Update POItems table.
    cnCon.CommitTrans
    
    '-- Check if all received.
        
    Load_MList
    txtOutput_Change
    Locate_Key
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
            Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim iCounter As Integer
    Dim c As New CDraw
    '
    ' Get some bitmaps over to the CDraw class so they
    ' can be used when drawing the list
    '
    
    '   Call CenterForm(frmSearch1)
   
    Set c.m_Picture1 = Picture1
    Set c.m_Picture2 = Picture1
    Set c.m_Picture3 = Picture1
    Set c.m_Picture4 = Picture1
    ' Assign the CDraw object to each column
    '
    MList1.Columns(0).PaintObject = c

    'MList1.Caption = "Search For () : " & Chr(13)
    'MList1.Columns(1).Heading = "Selection" & Chr(13) & "(選擇)"
    'MList1.Columns(2).Heading = "Level" & Chr(13) & "(級別)"
    'MList1.Columns(3).Heading = "Suffix" & Chr(13) & "Code (代碼)"
    'MList1.Columns(4).Heading = "Description" & Chr(13) & "(內容)"
    'MList1.Columns(5).Heading = "Selling" & Chr(13) & "Price (售價)"

    With MList1
        For iCounter = 1 To UBound(vHeadDataAry)
            .Columns(iCounter - 1).Heading = vHeadDataAry(iCounter, 1)
            '.Columns(iCounter - 1).Heading = GetLabel(vHeadDataAry(iCounter, 1), vHeadDataAry(iCounter, 1))
            '.Columns(iCounter - 1).Heading = GetFieldName(vHeadDataAry(iCounter, 1))
            .Columns(iCounter - 1).Width = vHeadDataAry(iCounter, 3)
        Next
    End With
    
    'L000401
    'Me.cmdQuickAdd.ToolTipText = GetToolTipNew(Me.Name, "cmdQuickAdd", "cmdQuickAdd")
    'If gsLangID <> "1" Then
    '    For Each vCtl In Me.Controls
    '        If TypeOf vCtl Is Label Then
    '            vCtl.Caption = GetLabelName(Me.Name, vCtl.Name)
    '        Else
    '            'If TypeOf vctl Is Frame Then
    '            '    vctl.Caption = GetLabelName(Me.Name, vctl.Name)
    '            'End If
    '        End If
    '    Next
    '
    '    Me.Caption = GetFormName(Me.Name)
    '
    '    'Tool Tip array start from 1
    '    vToolTip = Array("", "Locate", "Exit", "", "Refresh")
    '
    '    For iCounter = 1 To UBound(vToolTip)
    '        If Not vToolTip(iCounter) = "" Then
    '            Me.tbrProcess.Buttons.Item(iCounter).ToolTipText = GetToolTip(Me.Name, vToolTip(iCounter))
    '        End If
    '    Next
    '
    '    Me.cmdQuickAdd.ToolTipText = GetToolTip(Me.Name, "cmdQuickAdd")
    'End If
    'L000401 - end
    
    'If Not xLang(Me) Then
    '    MsgBox GetErrorMessage("E0001"), vbCritical + vbOKOnly, "Error"
    'End If
    If sInputType = "N" Then
    cmdQuickAdd.Visible = False
    Else
    cmdQuickAdd.Visible = True
    End If
    
    Load_MList
End Sub

Private Function GetID(str As String) As Long
    Dim i
    
    For i = 0 To MList1.ListCount - 1
        If UCase(GetSearchID(MList1.List(i), lSearchPos)) Like UCase(str) & "*" Then
            GetID = i
            Exit Function
        End If
    Next i
    GetID = -1
End Function

Private Function GetSearchID(str As String, lPos As Long) As String
    Dim i
    Dim oStr As String
    Dim iCounter As Integer
    
    iCounter = 0
    oStr = ""
    
    For i = 1 To Len(str)
    
        If Mid(str, i, 1) = MList1.ColDelimiter Then
            iCounter = iCounter + 1
            If iCounter = lPos Then
                GetSearchID = oStr
                Exit Function
            Else
                oStr = ""
            End If
        Else
            oStr = oStr & Mid(str, i, 1)
        End If
    Next i
    GetSearchID = oStr
End Function

Private Sub MList1_DblClick()
    Locate_Key
End Sub

Private Sub MList1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
            Unload Me
    End If
End Sub

Private Sub MList1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Locate_Key
        
    End If

End Sub

Private Sub txtOutput_Change()
    If txtOutput.Text = "" Then
        MList1.ListIndex = -1
    Else
        MList1.ListIndex = GetID(txtOutput.Text)
    End If
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Locate"
        Locate_Key
    Case "Exit"
        Unload frmPSearch
    Case "Refresh"
        Load_MList
End Select
End Sub

Private Sub Load_MList()
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String

    Set rsRcd = cnCon.Execute(sBindSQL)
    MList1.Clear
    
    If rsRcd.BOF Then Exit Sub
    
    rsRcd.MoveFirst
    Do While Not rsRcd.EOF
        sSQL = ""
        For iCounter = 1 To UBound(vHeadDataAry)
            If sSQL = "" Then
                sSQL = rsRcd(vHeadDataAry(iCounter, 2))
            Else
                sSQL = sSQL & "|" & rsRcd(vHeadDataAry(iCounter, 2))
            End If
        Next
        MList1.AddItem sSQL
        rsRcd.MoveNext
    Loop
    rsRcd.Close
End Sub

Private Sub Locate_Key()
    If MList1.ListIndex = -1 Then
        sMsg = "請於資料清單內選取至小一項"
        MsgBox sMsg, vbInformation + vbOKOnly, gsTitle
    Else
        oTextBox = GetSearchID(MList1.List(MList1.ListIndex), lKeyPos)
        Unload frmPSearch
    End If
End Sub


   
Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
            Unload Me
    End If
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        MList1.SetFocus
    End If
End Sub
