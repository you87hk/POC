VERSION 5.00
Begin VB.Form frmRegKey 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Reg Key"
   ClientHeight    =   3225
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5970
   ClipControls    =   0   'False
   Icon            =   "frmRegKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5970
   StartUpPosition =   2  '螢幕中央
   Tag             =   "Login"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   720
      Left            =   2880
      Picture         =   "frmRegKey.frx":030A
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   2160
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   720
      Left            =   1200
      Picture         =   "frmRegKey.frx":0614
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   2160
      Width           =   1620
   End
   Begin VB.TextBox txtWSID 
      Height          =   288
      IMEMode         =   3  '暫止
      Left            =   2145
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1395
      Width           =   2325
   End
   Begin VB.TextBox txtCPID 
      Height          =   288
      IMEMode         =   3  '暫止
      Left            =   2145
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblWSID 
      Caption         =   "Workstation ID:"
      Height          =   255
      Left            =   585
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   1410
      Width           =   1440
   End
   Begin VB.Label lblCPID 
      Caption         =   "Terminal ID:"
      Height          =   255
      Left            =   585
      TabIndex        =   5
      Tag             =   "&User Name:"
      Top             =   1020
      Width           =   1440
   End
End
Attribute VB_Name = "frmRegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

     txtCPID.Text = ""
     txtWSID.Text = ""
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'To Do - create test for correct password
    'check for correct password
    
    If Chk_txtWSID = False Then Exit Sub
    
    If Chk_txtCPID = False Then Exit Sub
    
 
    Call Write_Login_File
   ' Dim wssql
    
   ' wssql = "Insert into mstuser values ('TOM','TOM MOK','ADMIN', '" & Encrypt("TOMMOK") & "', 'NBASE','2000/01/01')"
   ' cnCon.Execute wssql
    
    Unload Me
    
   ' frmInfo.Show
 
End Sub

Private Sub Ini_Scr()

    Me.Caption = "Reg Key"
    lblCPID.Caption = "Terminal ID :"
    lblWSID.Caption = "WorkstationID : "
    cmdOK.Caption = "OK"
    cmdCancel.Caption = "Cancel"
    
        
End Sub

Private Sub Write_Login_File()

    Dim sBuffer As String
    Dim lSize As Long
    Dim WindowsPath As String
    Dim LoginFilePath As String
    Dim Mystring As String * 100
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetWindowsDirectory(sBuffer, lSize)
    If lSize > 0 Then
        WindowsPath = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        WindowsPath = vbNullString
    End If

    LoginFilePath = WindowsPath & "\NBASE.DAT"

    If Dir(LoginFilePath) <> "" Then
        Kill LoginFilePath
    End If
    Open LoginFilePath For Binary Access Write As #1
    wiCtr = 1
    
    Do While wiCtr < 3
        Select Case wiCtr
            Case 1
                Mystring = txtCPID.Text & txtWSID.Text & String(100 - Len(txtCPID.Text & txtWSID.Text), " ")
            Case 2
                Mystring = txtWSID.Text & String(100 - Len(txtWSID.Text), " ")
             End Select
        wiCtr = wiCtr + 1
        Put #1, , Encrypt(Mystring)
    Loop
    Close #1

End Sub


Private Sub txtWSID_GotFocus()
            
    Call FocusMe(txtWSID)
End Sub

Private Sub txtWSID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtWSID Then
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtWSID_LostFocus()
    Call FocusMe(txtWSID, True)
End Sub

Private Sub txtCPID_GotFocus()
    Call FocusMe(txtCPID)
End Sub

Private Sub txtCPID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtCPID Then
        txtWSID.SetFocus
        End If
    End If
End Sub

Private Sub txtCPID_LostFocus()
    Call FocusMe(txtCPID, True)
End Sub

Private Function Chk_txtCPID() As Boolean
    
    Chk_txtCPID = False
    
    If Trim(txtCPID.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCPID.SetFocus
        Exit Function
    End If
    
    Chk_txtCPID = True
End Function

Private Function Chk_txtWSID() As Boolean
    
    Chk_txtWSID = False
    
    If Trim(txtWSID.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtWSID.SetFocus
        Exit Function
    End If
    
    Chk_txtWSID = True
End Function

