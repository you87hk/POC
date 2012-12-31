Attribute VB_Name = "basMain"
Option Explicit

Public Sub Main()
'        Call Get_Login_File
'        frmOpn.ZOrder 1
       'frmOpn.Show vbModal
       'frmLogin.Show vbModal
        
        Call Read_Debug_Ini(App.Path & "\" & "HH_HZ.ini")
        gsUserID = "HHUSR"
        Call Connect_Database
        gsLangID = "2"
    
    gsDteFmt = "YMD"
    Call getHostLogin
    Call Get_Date_Fmt
    Call Get_Company_Info
    Call Get_Company_Default
    
    gsWhsCode = "1"
    gsCompID = "1"
    gsSystemDate = Format(Date, "yyyy/mm/dd")
    gsTitle = gsComNam
    
   ' nbMain.ZOrder 1
   ' nbMain.Show
   frmHHIM001.Show
   
End Sub

'-- Center form in center for the screen.
Public Sub CenterForm(oForm As Form)
    oForm.Top = (Screen.Height - oForm.Height) / 2
    oForm.Left = (Screen.Width - oForm.Width) / 2
End Sub

Public Sub Write_ErrLog_File(Mystring)

   ' Dim sBuffer As String
   ' Dim lSize As Long
    Dim WindowsPath As String
    Dim LoginFilePath As String
   ' Dim Mystring As String * 100
    
    On Error GoTo Err_Handler
    
    Exit Sub
    
    
  '  sBuffer = Space$(255)
  '  lSize = Len(sBuffer)
  '  Call GetWindowsDirectory(sBuffer, lSize)
  '  If lSize > 0 Then
  '      WindowsPath = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
  '  Else
  '      WindowsPath = vbNullString
  '  End If
    
    WindowsPath = "C:"
    LoginFilePath = WindowsPath & "\Errlog.TXT"

 '   If Dir(LoginFilePath) <> "" Then
 '       Kill LoginFilePath
 '   End If
    Open LoginFilePath For Append As #1
    
    Write #1, Now() & "-" & Mystring
    Close #1


    Exit Sub

Err_Handler:
    MsgBox Err.Description & " in Write_ErrLog_File"
    

End Sub

