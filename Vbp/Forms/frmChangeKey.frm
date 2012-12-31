VERSION 5.00
Begin VB.Form frmChangeKey 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   2385
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5775
   Icon            =   "frmChangeKey.frx":0000
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
         Picture         =   "frmChangeKey.frx":030A
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   735
         Left            =   120
         Picture         =   "frmChangeKey.frx":0614
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtKeyNo 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   5085
      End
      Begin VB.Label lblKeyNo 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmChangeKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property
Private mlKeyID    As Long
Private msKeyType  As String
Private mbResult  As Boolean
Private msNewKey  As String

Property Let KeyID(ByVal NewKeyID As Long)

   mlKeyID = NewKeyID

End Property

Property Let KeyType(ByVal NewKeyType As String)

   msKeyType = NewKeyType

End Property



Property Get Result() As Boolean

   Result = mbResult
   
End Property

Property Get NewKey() As String

   NewKey = msNewKey
   
End Property

Private Sub btnCancel_Click()
    mbResult = False
    Unload Me
End Sub

Private Sub btnOK_Click()

If Chk_txtKeyNo() = False Then Exit Sub

    mbResult = True
    msNewKey = txtKeyNo.Text
    Unload Me

End Sub

Private Sub Form_Load()
 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
  
    mbResult = False
    msNewKey = ""
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 

 Set waScrItm = Nothing
 Set frmChangeKey = Nothing
    
End Sub

Private Sub txtKeyNo_GotFocus()
    FocusMe txtKeyNo
End Sub

Private Sub txtKeyNo_KeyPress(KeyAscii As Integer)
     Call chk_InpLenA(txtKeyNo, 30, KeyAscii, True)
  
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtKeyNo() = False Then Exit Sub
        
        btnOK.SetFocus
        
  End If
    
End Sub

Private Sub txtKeyNo_LostFocus()
    FocusMe txtKeyNo, True
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "ChangeKey"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    lblKeyNo.Caption = Get_Caption(waScrItm, "KeyNo")
    btnOK.Caption = Get_Caption(waScrItm, "OK")
    btnCancel.Caption = Get_Caption(waScrItm, "CANCEL")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtKeyNo() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsMsg As String
    
    Chk_txtKeyNo = False
    
    If Trim(txtKeyNo.Text) = "" Then
        wsMsg = "物料編號由系統設定!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtKeyNo = Get_ItemNo(msKeyType)
        Chk_txtKeyNo = True
        Exit Function
    End If
    
    wsSQL = "SELECT * FROM mstItem WHERE ItmCode = '" & Set_Quote(txtKeyNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        
            wsMsg = "物料編號已存在!"
            MsgBox wsMsg, vbOKOnly, gsTitle
            txtKeyNo.SetFocus
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    Chk_txtKeyNo = True

End Function


