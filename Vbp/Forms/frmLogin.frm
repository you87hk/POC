VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Login"
   ClientHeight    =   3375
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   3375
   ScaleWidth      =   7650
   StartUpPosition =   2  '螢幕中央
   Tag             =   "Login"
   Begin VB.ComboBox cboLangID 
      Height          =   300
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2250
   End
   Begin VB.ComboBox cboCompany 
      Height          =   300
      Left            =   3465
      TabIndex        =   0
      Top             =   630
      Width           =   4056
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   720
      Left            =   5040
      Picture         =   "frmLogin.frx":2581
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Tag             =   "Cancel"
      Top             =   2400
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   720
      Left            =   3360
      Picture         =   "frmLogin.frx":288B
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   2400
      Width           =   1620
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      IMEMode         =   3  '暫止
      Left            =   3465
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1395
      Width           =   2325
   End
   Begin VB.TextBox txtUserID 
      Height          =   288
      Left            =   3465
      TabIndex        =   1
      Top             =   990
      Width           =   2325
   End
   Begin VB.Label lblLanguage 
      BackStyle       =   0  '透明
      Caption         =   "&Language:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   9
      Tag             =   "&User Name:"
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  '透明
      Caption         =   "&Company:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   8
      Tag             =   "&User Name:"
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  '透明
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   6
      Tag             =   "&Password:"
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  '透明
      Caption         =   "&User ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      TabIndex        =   7
      Tag             =   "&User Name:"
      Top             =   1020
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wsWSID As String
Private wsCPID As String
Private wsINIFile As String



Private Sub cboCompany_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtUserID.SetFocus
    End If
End Sub



Private Sub cboLangID_GotFocus()
FocusMe cboLangID
End Sub

Private Sub cboLangID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cmdOK.SetFocus
    End If
End Sub

Private Sub cboLangID_LostFocus()
    FocusMe cboLangID, True
End Sub

Private Sub Form_Load()

    Dim sBuffer As String
    Dim lSize As Long
    Dim wiCtr As Long
    
    
   ' If App.PrevInstance = True Then
   '     End
   ' End If
    
   ' frmSplash.Show
   ' frmSplash.Refresh
    
   ' For wiCtr = 0 To 200000
   '     DoEvents
   ' Next
    
   ' Unload frmSplash
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserID.Text = Left$(sBuffer, lSize)
    Else
        txtUserID.Text = vbNullString
    End If
    
    Call Ini_Scr
    
    Call GetCompanyList
    Call GetLanugaeList
    
End Sub

Private Sub cmdCancel_Click()
    End
End Sub


Private Sub cmdOK_Click()
    'To Do - create test for correct password
    'check for correct password
    
    Dim Chk_Login_Result As Integer
    
    Dim wiIndex As Integer
    
    If cboCompany.ListIndex = -1 Then
        For wiIndex = 0 To cboCompany.ListCount - 1
            If Trim(UCase(cboCompany.Text)) = UCase(cboCompany.List(wiIndex)) Then
                cboCompany.ListIndex = wiIndex
                Exit For
            End If
        Next
    End If
    
    Call Get_Selected_INI(cboCompany.ListIndex)
    
    If Dir(wsINIFile) = "" Or Trim(wsINIFile) = "" Then
        MsgBox "Can't Find ini File!"
        End
    Else
        Call Read_Debug_Ini(wsINIFile)
    End If
    
   If cboLangID.ListIndex = -1 Then
        For wiIndex = 0 To cboLangID.ListCount - 1
            If Trim(UCase(cboLangID.Text)) = UCase(cboLangID.List(wiIndex)) Then
                cboLangID.ListIndex = wiIndex
                Exit For
            End If
        Next
    End If
    
    Call Get_Selected_Lang(cboLangID.ListIndex)
    
    If Connect_Database = False Then End
    
    Chk_Login_Result = Chk_Login(txtUserID.Text, txtPassword.Text)
    
    Select Case Chk_Login_Result
      '  Case 0
      '      Me.Hide
        Case 1
            gsMsg = "沒有此用戶!"
            MsgBox gsMsg, vbInformation + vbOKOnly
            txtUserID.SetFocus
            GoTo Login_Err
        Case 2
            gsMsg = "密碼錯誤!"
            MsgBox gsMsg, vbInformation + vbOKOnly
            txtPassword.SetFocus
            GoTo Login_Err
    End Select
    
 '   If UCase(txtUserID.Text) = "NBASE" Then
 '       gsWorkStationID = "01"
 '   Else
 '        Call GetSystemData
 '   End If
    
 '   Call Get_Company_Info
        
  '  Call Write_Login_File
    
   ' frmInfo.Show
    Unload Me
    Exit Sub
    
Login_Err:
    Call Disconnect_Database
    
End Sub

Private Sub Ini_Scr()

    Me.Caption = "Log In"
    lblCompany.Caption = "Company :"
    lblUser.Caption = "User Name :"
    lblPassword.Caption = "Password: "
    lblLanguage.Caption = "Language :"
    cmdOK.Caption = "OK"
    cmdCancel.Caption = "Cancel"
    
        
End Sub
Private Sub GetCompanyList()

    Dim sBuffer As String
    Dim lSize As Long
    Dim SystemPath As String
    Dim CompanyEntries As String
         
    On Error GoTo Err_GetCompanyList
    
    SystemPath = App.Path
      
    If Dir(SystemPath & "\COMPANY.LST") = "" Then
        MsgBox LoadResString(113)
        End
    Else
        Open SystemPath & "\COMPANY.LST" For Input As #1
        cboCompany.Clear
        Do While Not EOF(1)
            Input #1, CompanyEntries
            If InStr(1, CompanyEntries, ";") > 0 Then
               cboCompany.AddItem Left(CompanyEntries, InStr(1, CompanyEntries, ";") - 1)
            End If
        Loop
        Close #1
        cboCompany.ListIndex = 0
    End If
    
    Exit Sub
    
Err_GetCompanyList:

   gsMsg = "找不到公司清單!"
   MsgBox gsMsg, vbInformation + vbOKOnly
   
End Sub

Private Sub GetLanugaeList()

    Dim sBuffer As String
    Dim lSize As Long
    Dim SystemPath As String
    Dim LangEntries As String
         
    On Error GoTo Err_GetLanugaeList
    
    SystemPath = App.Path
      
    If Dir(SystemPath & "\LANG.LST") = "" Then
        MsgBox LoadResString(113)
        End
    Else
        Open SystemPath & "\LANG.LST" For Input As #1
        cboLangID.Clear
        Do While Not EOF(1)
            Input #1, LangEntries
            If InStr(1, LangEntries, ";") > 0 Then
               cboLangID.AddItem Left(LangEntries, InStr(1, LangEntries, ";") - 1)
            End If
        Loop
        Close #1
        cboLangID.ListIndex = 0
    End If
    
    Exit Sub
    
Err_GetLanugaeList:

   gsMsg = "找不到語種清單!"
   MsgBox gsMsg, vbInformation + vbOKOnly
   
End Sub

Private Function Get_Selected_INI(inListindex As Integer) As String
    
    Dim compLine As String
    Dim Counter As Integer
    Dim EndofFirstPart As Integer
    
    Counter = 0
    Get_Selected_INI = ""
    
    Open App.Path & "\COMPANY.LST" For Input As #1
    Do While Not EOF(1) And Counter <= inListindex
        Line Input #1, compLine
        If Counter = inListindex Then
            EndofFirstPart = InStr(1, compLine, ";")
            Get_Selected_INI = App.Path & "\" & Mid(compLine, EndofFirstPart + 1)
        End If
        Counter = Counter + 1
    Loop
    Close #1
    wsINIFile = Get_Selected_INI

End Function

Private Function Get_Selected_Lang(inListindex As Integer) As String
    
    Dim compLine As String
    Dim Counter As Integer
    Dim EndofFirstPart As Integer
    
    Counter = 0
    Get_Selected_Lang = ""
    
    Open App.Path & "\LANG.LST" For Input As #1
    Do While Not EOF(1) And Counter <= inListindex
        Line Input #1, compLine
        If Counter = inListindex Then
            EndofFirstPart = InStr(1, compLine, ";")
            Get_Selected_Lang = Mid(compLine, EndofFirstPart + 1)
        End If
        Counter = Counter + 1
    Loop
    Close #1
    gsLangID = Get_Selected_Lang

End Function

Private Function Chk_Login(inUser, inPassword) As Integer

    Dim rsLogin As New ADODB.Recordset
    Dim Criteria As String
    
    Chk_Login = 0
    
    Criteria = "SELECT USRPWD, USRNAME FROM MSTUSER WHERE USRCODE = '" & Set_Quote(inUser) & "' "

    rsLogin.Open Criteria, cnCon, adOpenStatic, adLockOptimistic

    If rsLogin.RecordCount > 0 Then
        If Encrypt(rsLogin("USRPWD")) <> UCase(inPassword) Then
            Chk_Login = 2
        Else
            gsUserID = inUser
        End If
    Else
        If UCase(inUser) = "NBASE" And UCase(inPassword) = "NBTEL" Then
            gsUserID = UCase(inUser)
        Else
            Chk_Login = 1
        End If
    End If
    rsLogin.Close
    
End Function



Private Sub GetSystemData()
        
    
    Dim sBuffer As String
    Dim lSize As Long
    Dim WindowsPath As String
    Dim LoginFilePath As String
    Dim Mystring As String * 100
    Dim LineCount As Integer
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetWindowsDirectory(sBuffer, lSize)
    If lSize > 0 Then
        WindowsPath = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        WindowsPath = vbNullString
    End If
        
      
    If Dir(WindowsPath & "\NBASE.DAT") = "" Then
        MsgBox "No System File!"
        End
    Else
        Open WindowsPath & "\NBASE.DAT" For Binary Access Read As #1
        LineCount = 1
        Do While Not EOF(1) And LineCount < 3
            Get #1, , Mystring
            Select Case LineCount
                Case 1
                    wsCPID = Trim(Encrypt(Mystring))
                Case 2
                    wsWSID = Trim(Encrypt(Mystring))
            End Select
            LineCount = LineCount + 1
        Loop
        Close #1
    End If
    
    If UCase(Right(wsCPID, 2)) <> UCase(wsWSID) Then
        MsgBox "Non-authorize User!"
        End
    End If
    
    gsWorkStationID = wsWSID
        
End Sub



Private Sub txtPassword_GotFocus()
    Call FocusMe(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
       cmdOK_Click
    End If
End Sub

Private Sub txtPassword_LostFocus()
    Call FocusMe(txtPassword, True)
End Sub

Private Sub txtUserID_GotFocus()
    Call FocusMe(txtUserID)
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtPassword.SetFocus
    End If
End Sub

Private Sub txtUserID_LostFocus()
    Call FocusMe(txtUserID, True)
End Sub
