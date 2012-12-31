VERSION 5.00
Begin VB.Form frmKeyInput 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   2385
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5775
   Icon            =   "frmKeyInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Begin VB.TextBox txtKeyNo 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lblKeyNo 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmKeyInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ctlKey As Control

Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property
Private msTableID    As String
Private msTableType  As String
Private msTableKey   As String
Private msKeyLen     As Integer



Property Get TableID() As String

   TableID = msTableID
   
End Property
Property Let TableID(ByVal NewTableID As String)

   msTableID = NewTableID

End Property


Property Get TableKey() As String

   TableKey = msTableKey
   
End Property
Property Let TableKey(ByVal NewTableKey As String)

   msTableKey = NewTableKey

End Property


Property Get TableType() As String

   TableType = msTableType
   
End Property
Property Let TableType(ByVal NewTableType As String)

   msTableType = NewTableType

End Property

Property Get KeyLen() As Integer

   KeyLen = msKeyLen
   
End Property
Property Let KeyLen(ByVal NewKeyLen As Integer)

   msKeyLen = NewKeyLen

End Property

Private Sub Form_Load()
 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
  
  
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 
 If Chk_txtKeyNo() = False Then
    
    Cancel = True
    Exit Sub
        
 End If
 
 Set waScrItm = Nothing
 Set frmKeyInput = Nothing
    
End Sub

Private Sub txtKeyNo_GotFocus()
    FocusMe txtKeyNo
End Sub

Private Sub txtKeyNo_KeyPress(KeyAscii As Integer)
     Call chk_InpLenA(txtKeyNo, KeyLen, KeyAscii, True)
  
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtKeyNo() = False Then Exit Sub
        
        ctlKey = txtKeyNo
        Unload Me
            
  End If
    
End Sub

Private Sub txtKeyNo_LostFocus()
    FocusMe txtKeyNo, True
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "KeyInput"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    lblKeyNo.Caption = Get_Caption(waScrItm, "KeyNo")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtKeyNo() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsMsg As String
    
    Chk_txtKeyNo = False
    
    If Trim(txtKeyNo.Text) = "" And Chk_AutoGen(TableType) = "N" Then
        wsMsg = "Key No Must Input!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtKeyNo.SetFocus
        Exit Function
    End If
    
    wsSQL = "SELECT * FROM " & TableID & " WHERE " & TableKey & " = '" & Set_Quote(txtKeyNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        
            wsMsg = "Key Already Exist!"
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
