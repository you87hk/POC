VERSION 5.00
Begin VB.Form frmHH 
   Caption         =   "Receive from HH"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5670
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox HHList 
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnReceive 
      Caption         =   "Receive"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wsHHPath As String



Private Sub btnImport_Click()
    Call cmdImport
End Sub

Private Sub btnLoad_Click()

    Dim MyFile

    
    HHList.Clear
    
    MyFile = Dir(wsHHPath, vbNormal)
    Do While MyFile <> ""
    
    HHList.AddItem MyFile
    
    MyFile = Dir
    Loop
    
    

End Sub

Private Sub cmdImport()
Dim sFullFile As String
Dim sFileName As String
Dim sExt As String
Dim SDOCNO As String
Dim i As Integer
Dim bPost As Boolean
Dim relUpdFlg As String

Dim sfile As String
Dim wsDteTim As String


     
  MousePointer = vbHourglass
  
   wsDteTim = Change_SQLDate(Now)
  gsMsg = ""

For i = 0 To HHList.ListCount - 1 'hidden ListBox
   ' sFullFile = HHList.List(i)
   ' sFileName = Right(sFullFile, Len(sFullFile) - InStrRev(sFullFile, "\"))
   
    sFileName = HHList.List(i)
    sFullFile = wsHHPath & HHList.List(i)
    sExt = Right(sFileName, Len(sFileName) - InStr(sFileName, "."))

    SDOCNO = Left(sFileName, InStr(sFileName, ".") - 1)
    
    
    If UCase(sExt) <> "BAK" And UCase(sExt) <> "FLD" Then
    
    If Chk_HHNo(SDOCNO, relUpdFlg) = True Then
        If relUpdFlg = "N" Then
        
            gsMsg = SDOCNO & "已匯入但未更新, 你是否確認要覆寫此文件?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
            bPost = False
            Else
            bPost = True
            End If
        
        Else
            gsMsg = SDOCNO & " 已匯入並更新!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            bPost = False
           ' Name sFullFile As Left(sFullFile, Len(sFullFile) - 3) & "BAK"
        End If
    Else
        bPost = True
    End If
    
    If bPost Then
        If ImportFromHH(gsUserID, wsDteTim, sFullFile) = True Then
        gsMsg = gsMsg & IIf(gsMsg = "", SDOCNO, Chr(10) & Chr(13) & SDOCNO)
        End If
    End If
    
    End If

    Name sFullFile As wsHHPath & "backup\" & sFileName
    
Next i


 
 gsMsg = "匯入完成!"
 MsgBox gsMsg, vbOKOnly, gsTitle
 

 MousePointer = vbDefault

    
    
End Sub

Private Sub btnReceive_Click()


   MousePointer = vbHourglass
    
  Call ReceiveFromHH(wsHHPath)
  Sleep (1000)
  
MousePointer = vbNormal
  

End Sub

Private Sub Ini_Scr()


    If Trim(gsHHPath) <> "" Then
    wsHHPath = gsHHPath + "receive\"
    Else
    wsHHPath = App.Path + "receive\"
    End If

End Sub

Private Sub Form_Load()

    Call Ini_Scr

End Sub
