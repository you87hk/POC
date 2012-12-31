VERSION 5.00
Begin VB.Form frmPrintStop 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Printing"
   ClientHeight    =   810
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   1965
   ControlBox      =   0   'False
   Icon            =   "frmPrintStop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   1965
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrintStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Declare the constant value
'Option




Private Sub btnClose_Click()
   Unload Me
End Sub

Private Sub Form_Activate()

   Me.MousePointer = vbHourglass

   Ini_Scr
   Ini_Caption

   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Me.MousePointer = vbHourglass
 

   Me.MousePointer = vbDefault

   

   
End Sub


Private Sub Form_LostFocus()

   'Unload Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'SaveUserDefault
   Set frmPrintBar = Nothing
   

   
End Sub


Private Sub Ini_Scr()

   
   Me.Visible = True
   


End Sub





Private Sub Ini_Caption()


   
End Sub

