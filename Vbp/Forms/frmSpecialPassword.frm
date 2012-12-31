VERSION 5.00
Begin VB.Form frmSpecialPassword 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   2385
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5775
   Icon            =   "frmSpecialPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5775
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraHeader 
      Height          =   2040
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   735
         Left            =   1560
         Picture         =   "frmSpecialPassword.frx":030A
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   735
         Left            =   120
         Picture         =   "frmSpecialPassword.frx":0614
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   5085
      End
      Begin VB.Label lblPassword 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmSpecialPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property
Private mbResult  As Boolean
Private msDocNo As String


Property Let DocNo(ByVal InDocNo As String)

   msDocNo = InDocNo

End Property

Property Get Result() As Boolean

   Result = mbResult
   
End Property

Private Sub btnCancel_Click()
    mbResult = False
    Unload Me
End Sub

Private Sub btnOK_Click()

If Chk_txtPassword() = False Then Exit Sub

    mbResult = True
    Unload Me

End Sub

Private Sub Form_Load()
 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
  
    mbResult = False
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 

 Set waScrItm = Nothing
 Set frmSpecialPassword = Nothing
    
End Sub

Private Sub txtPassword_GotFocus()
    FocusMe txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    ' Call chk_InpLenA(txtPassword, 30, KeyAscii, True)
  
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPassword() = False Then Exit Sub
        
        btnOK.SetFocus
        
  End If
    
End Sub

Private Sub txtPassword_LostFocus()
    FocusMe txtPassword, True
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    txtPassword.PasswordChar = "*"
    wsFormID = "SpecPwd"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    lblPassword.Caption = Get_Caption(waScrItm, "Password") + " (" & msDocNo & ")"
    btnOK.Caption = Get_Caption(waScrItm, "OK")
    btnCancel.Caption = Get_Caption(waScrItm, "CANCEL")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtPassword() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsMsg As String
    
    Chk_txtPassword = False
    
    If Trim(txtPassword.Text) = "" Then
        wsMsg = "請輸入特定密碼!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Chk_txtPassword = False
        Exit Function
    End If
    
    
    If Chk_SpecialPassword(txtPassword.Text) = False Then
    wsMsg = "密碼錯誤!"
    MsgBox wsMsg, vbOKOnly, gsTitle
    txtPassword.Text = ""
    Chk_txtPassword = False
    Else
    Chk_txtPassword = True
    End If
    
    

End Function


