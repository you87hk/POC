VERSION 5.00
Begin VB.Form frmCHGPWD 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Chang Pasword"
   ClientHeight    =   3375
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6600
   Icon            =   "frmChgPwd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6600
   StartUpPosition =   2  '螢幕中央
   Tag             =   "Login"
   Begin VB.TextBox txtNewPassword 
      Height          =   288
      IMEMode         =   3  '暫止
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   720
      Left            =   3480
      Picture         =   "frmChgPwd.frx":030A
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   2280
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   720
      Left            =   1800
      Picture         =   "frmChgPwd.frx":0614
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   2280
      Width           =   1620
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      IMEMode         =   3  '暫止
      Left            =   2745
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1155
      Width           =   2325
   End
   Begin VB.Label lblDspUserID 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   720
      Width           =   2265
   End
   Begin VB.Label lblNewPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Tag             =   "&Password:"
      Top             =   1575
      Width           =   1080
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   1545
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   1170
      Width           =   1080
   End
   Begin VB.Label lblUser 
      Caption         =   "&User ID:"
      Height          =   255
      Left            =   1545
      TabIndex        =   5
      Tag             =   "&User Name:"
      Top             =   780
      Width           =   1080
   End
End
Attribute VB_Name = "frmCHGPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wsFormCaption As String
Private waScrItm As New XArrayDB

 

Private Const wsKeyType = "MstAccountType"
Private wsUsrId As String
Private wsFormID As String
Private wsConnTime As String
Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblUser.Caption = Get_Caption(waScrItm, "USERID")
    lblPassword.Caption = Get_Caption(waScrItm, "PASSWORD")
    lblNewPassword.Caption = Get_Caption(waScrItm, "NEWPASSWORD")
    
    cmdOK.Caption = Get_Caption(waScrItm, "OK")
    cmdCancel.Caption = Get_Caption(waScrItm, "CANCEL")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub
Private Sub IniForm()
    Me.KeyPreview = True
 
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "CHGPWD"
 
    
End Sub

Private Sub Ini_Scr()

    lblDspUserID.Caption = gsUserID
    txtPassword.Text = ""
    txtNewPassword.Text = ""
    
    Me.Caption = wsFormCaption
        
End Sub
Private Sub Form_Load()

    MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
  
  
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
    'To Do - create test for correct password
    'check for correct password
    
    Dim Chk_Password_Result As Integer
    
   
    Chk_Password_Result = Chk_Password(lblDspUserID.Caption, txtPassword.Text)
    
    Select Case Chk_Password_Result
      '  Case 0
      '      Me.Hide
       
        Case 1
            gsMsg = "密碼錯誤!"
            MsgBox gsMsg, vbInformation + vbOKOnly
            txtPassword.SetFocus
            Exit Sub
    End Select
    

    If Chk_txtNewPassword = False Then Exit Sub

    If cmdSave Then Unload Me

    

    
End Sub


Private Function Chk_Password(inUser, inPassword) As Integer

    Dim rsLogin As New ADODB.Recordset
    Dim Criteria As String
    
    Chk_Password = 1
    
    Criteria = "SELECT USRPWD, USRNAME FROM MSTUSER WHERE USRCODE = '" & Set_Quote(inUser) & "' "

    rsLogin.Open Criteria, cnCon, adOpenStatic, adLockOptimistic

    If rsLogin.RecordCount > 0 Then
        If Encrypt(rsLogin("USRPWD")) <> UCase(inPassword) Then
            Chk_Password = 1
            Exit Function
        End If
    Else
            Exit Function
    End If
    rsLogin.Close
    
    Chk_Password = 0
End Function





Private Sub txtPassword_GotFocus()
    Call FocusMe(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtNewPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_LostFocus()
    Call FocusMe(txtPassword, True)
End Sub

Private Sub txtNewPassword_GotFocus()
    Call FocusMe(txtNewPassword)
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtNewPassword Then
        cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtNewPassword_LostFocus()
    Call FocusMe(txtNewPassword, True)
End Sub


Private Function Chk_txtNewPassword() As Boolean
    Dim wsStatus As String

    Chk_txtNewPassword = False
    
        If Trim(txtNewPassword.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtNewPassword.SetFocus
            Exit Function
        End If
    
        
    Chk_txtNewPassword = True
End Function
Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    
   
    cmdSave = False
   
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_CHGPWD"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, 2)
    Call SetSPPara(adcmdSave, 2, lblDspUserID.Caption)
    Call SetSPPara(adcmdSave, 3, Encrypt(UCase(Set_Quote(txtNewPassword.Text))))
    Call SetSPPara(adcmdSave, 4, gsUserID)
    Call SetSPPara(adcmdSave, 5, wsGenDte)
    
    adcmdSave.Execute
    
    cnCon.CommitTrans
    
    
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

