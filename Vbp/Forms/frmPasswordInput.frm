VERSION 5.00
Begin VB.Form frmPasswordInput 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   2850
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5775
   Icon            =   "frmPasswordInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5775
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraHeader 
      Height          =   2640
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5475
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   720
         Left            =   720
         Picture         =   "frmPasswordInput.frx":030A
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Tag             =   "OK"
         Top             =   1800
         Width           =   1620
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   720
         Left            =   2400
         Picture         =   "frmPasswordInput.frx":0614
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Tag             =   "Cancel"
         Top             =   1800
         Width           =   1620
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   1  '靠右對齊
         Height          =   288
         IMEMode         =   3  '暫止
         Left            =   600
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label lblDspDesc 
         Caption         =   "NEW KEY:"
         Height          =   720
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblPassword 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmPasswordInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property
Private wsDesc As String
Private wbResult  As Boolean




Property Let InDesc(ByVal InputDesc As String)

   wsDesc = InputDesc

End Property


Property Get pResult() As Boolean

   pResult = wbResult
   
End Property



Private Sub cmdCancel_Click()
        Unload Me
End Sub

Private Sub cmdOK_Click()


        If Chk_txtPassword = True Then
            wbResult = True
            Unload Me
        End If
        
End Sub

Private Sub Form_Load()
 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Scr
  
    MousePointer = vbDefault

End Sub
Private Sub Ini_Scr()

    Dim MyControl As Control
    
      For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
            Case "ComboBox"
                MyControl.Clear
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


    wbResult = False
    
    Call SetPasswordChar(txtPassword, "*")
    
    lblDspDesc.Caption = wsDesc
    
    FocusMe txtPassword

End Sub


Private Sub Form_Unload(Cancel As Integer)
 

 Set waScrItm = Nothing
 
    
End Sub

Private Sub txtPassword_GotFocus()
    FocusMe txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPassword = True Then
            wbResult = True
            Unload Me
        End If
            
  End If
    
End Sub

Private Sub txtPassword_LostFocus()
    FocusMe txtPassword, True
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "PassInput"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    lblPassword.Caption = Get_Caption(waScrItm, "PASSWORD")
    cmdCancel.Caption = Get_Caption(waScrItm, "CANCEL")
    cmdOK.Caption = Get_Caption(waScrItm, "OK")
    
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtPassword() As Boolean
                      
 Chk_txtPassword = False
                      
 If Chk_SpecialPassword(txtPassword.Text) = False Then
        
        gsMsg = "密碼錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPassword.SetFocus
        Exit Function
 End If

Chk_txtPassword = True

End Function

  
