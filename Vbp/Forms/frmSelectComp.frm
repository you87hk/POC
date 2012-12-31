VERSION 5.00
Begin VB.Form frmSelectComp 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "New Key"
   ClientHeight    =   2385
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5250
   Icon            =   "frmSelectComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5250
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
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
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Option1"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
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
Attribute VB_Name = "frmSelectComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ctlKey As String
 Public wsCmpCode As String
Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property

Private Sub Form_Load()

 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
    
    wsCmpCode = Get_CompanyFlag("CMPCODE")
    
    Select Case wsCmpCode
    Case "FR"
        OptComp(0).Value = True
    Case "CF"
        OptComp(1).Value = True
    Case "CY"
        OptComp(2).Value = True
    End Select
        
    
 
  
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 
 If Chk_txtKeyNo() = False Then
    
    Cancel = True
    Exit Sub
    
 Else
 ctlKey = UCase(wsCmpCode)
        
 End If
 
 Set waScrItm = Nothing
 Set frmSelectComp = Nothing
    
End Sub



Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "SelectComp"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    'lblKeyNo.Caption = Get_Caption(waScrItm, "KeyNo")
    Me.Caption = "Copy Quotation"
    
    fraSelect.Caption = "Company Select"
    
    OptComp(0).Caption = "Fair Rack (½÷°ì)"
    OptComp(1).Caption = "Chung Fai (©¾½÷)"
    OptComp(2).Caption = "Cyber Fortune (¬ÕºÖ)"
     
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtKeyNo() As Boolean
    
    Dim wsMsg As String
    
    Chk_txtKeyNo = False
    
    If OptComp(0).Value Then
        wsCmpCode = "FR"
    ElseIf OptComp(1).Value Then
        wsCmpCode = "CF"
    Else
        wsCmpCode = "CY"
    End If
    
    
    If Trim(wsCmpCode) = "" Then
        wsMsg = "Company Code Must Input!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    If UCase(wsCmpCode) <> "FR" And UCase(wsCmpCode) <> "CF" And UCase(wsCmpCode) <> "CY" Then
            
            wsMsg = "Company Coode Not Exist!"
            MsgBox wsMsg, vbOKOnly, gsTitle
            Exit Function
        
    End If
    
    
    Chk_txtKeyNo = True

End Function
