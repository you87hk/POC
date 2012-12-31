VERSION 5.00
Begin VB.Form frmDocRemark 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   5850
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   5730
   Icon            =   "frmDocRemark.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5730
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraHeader 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.TextBox txtRmk 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4875
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   1
         Top             =   480
         Width           =   5115
      End
      Begin VB.Label lblRmk 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmDocRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private wsFormID As String
Private waScrItm As New XArrayDB

'variable for new property
Private msRmkID    As String
Private msRmkType  As String

Private wbExit    As Boolean
Private wsOldRmk    As String




Property Get RmkID() As String

   RmkID = msRmkID
   
End Property

Property Let RmkID(ByVal NewRmkID As String)

   msRmkID = NewRmkID

End Property
Property Let RmkType(ByVal NewRmkType As String)

   msRmkType = NewRmkType

End Property



Private Sub Form_Load()
 
 MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Scr
  
  
    MousePointer = vbDefault

End Sub



Private Sub Form_Unload(Cancel As Integer)
 
 If wbExit = False Then
 
    Call cmdSave
    wbExit = True
    Cancel = True
    Me.Hide
    Exit Sub
        
 End If
 
 Set waScrItm = Nothing
 Set frmDocRemark = Nothing
    
End Sub

Private Sub txtRmk_GotFocus()
'    FocusMe txtRmk
End Sub

Private Sub txtRmk_KeyPress(KeyAscii As Integer)
 'Call chk_InpLen(txtRmk, KeyLen, KeyAscii)
  
  
 '   If Len(txtRmk.Text) Mod 50 = 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
 '       KeyAscii = vbKeyReturn
 '   End If
  
  
'  If KeyAscii = vbKeyReturn Then
'        KeyAscii = vbDefault
        
'        If Chk_txtRmk() = False Then Exit Sub
        
       
            
'  End If
    
End Sub

Private Sub txtRmk_LostFocus()
    FocusMe txtRmk, True
End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "DocRemark"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    
    lblRmk.Caption = Get_Caption(waScrItm, "Rmk")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_txtRmk() As Boolean
    
    Dim wsMsg As String
    
    Chk_txtRmk = False
    
    If Trim(txtRmk.Text) = "" Then
        wsMsg = "Remark Must Input!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtRmk.SetFocus
        Exit Function
    End If
    
    Chk_txtRmk = True

End Function

Private Sub Ini_Scr()

    
    
    wbExit = False
    
    Call LoadRecord
    
    
End Sub

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT DRREMARK "
    wsSQL = wsSQL + "FROM MSTDOCREMARK "
    wsSQL = wsSQL + "WHERE DRID = " & msRmkID
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        Me.txtRmk = ReadRs(rsRcd, "DRREMARK")
        wsOldRmk = txtRmk.Text
        LoadRecord = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function


Private Function cmdSave() As Boolean
    Dim wsGenDte As String

    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If Trim(txtRmk.Text) = "" Then
            msRmkID = "0"
            MousePointer = vbDefault
            Exit Function
    End If
    
    If wsOldRmk = txtRmk.Text Then
            MousePointer = vbDefault
            Exit Function
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_DR001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, msRmkID)
    Call SetSPPara(adcmdSave, 2, msRmkType)
    Call SetSPPara(adcmdSave, 3, txtRmk.Text)
    Call SetSPPara(adcmdSave, 4, gsUserID)
    Call SetSPPara(adcmdSave, 5, wsGenDte)
    adcmdSave.Execute
    msRmkID = GetSPPara(adcmdSave, 6)
    
    cnCon.CommitTrans
    
    If Trim(msRmkID) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - EXC001!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
        gsMsg = "已成功儲存!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    End If
    
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            
        Case vbKeyEscape
            Unload Me
    End Select
End Sub
