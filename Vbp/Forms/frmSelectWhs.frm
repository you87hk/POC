VERSION 5.00
Begin VB.Form frmSelectWhs 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   2385
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5250
   Icon            =   "frmSelectWhs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5250
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraSelect 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.OptionButton OptComp 
         Caption         =   "Option1"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSelectWhs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ctlKey As String
 Public wsWhsCode As String
Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property

Private Sub Form_Load()

 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
    
    wsWhsCode = Get_TableInfo("SYSWSINFO", "WSID = '01'", "WSWHSCODE")
    
    Select Case wsWhsCode
    Case "CF-HK"
        OptComp(0).Value = True
    Case "CF-HZ"
        OptComp(1).Value = True
    Case Else
        OptComp(0).Value = True
    End Select
        
    
 
  
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 
 If Chk_txtKeyNo() = False Then
    
    Cancel = True
    Exit Sub
    
 Else
 ctlKey = UCase(wsWhsCode)
        
 End If
 
 Set waScrItm = Nothing
 Set frmSelectWhs = Nothing
    
End Sub



Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "SelectComp"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    'lblKeyNo.Caption = Get_Caption(waScrItm, "KeyNo")
    Me.Caption = "Select Warehouse"
    
    fraSelect.Caption = "Where is the location (倉庫位置)"
    
    OptComp(0).Caption = "Hong Kong(香港倉)"
    OptComp(1).Caption = "H.Z. (鶴山倉)"
     
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtKeyNo() As Boolean
    
    Dim wsMsg As String
    
    Chk_txtKeyNo = False
    
    If OptComp(0).Value Then
        wsWhsCode = "CF-HK"
    ElseIf OptComp(1).Value Then
        wsWhsCode = "CF-HZ"
    Else
        wsWhsCode = "CF-HZ"
    End If
    
    
    If Trim(wsWhsCode) = "" Then
        wsMsg = "Warehouse Code Must Input!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    If UCase(wsWhsCode) <> "CF-HK" And UCase(wsWhsCode) <> "CF-HZ" Then
            
            wsMsg = "Warehouse Coode Not Exist!"
            MsgBox wsMsg, vbOKOnly, gsTitle
            Exit Function
        
    End If
    
    
    Chk_txtKeyNo = True

End Function
