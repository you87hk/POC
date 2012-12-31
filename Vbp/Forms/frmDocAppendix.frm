VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDocAppendix 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   9045
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   12840
   Icon            =   "frmDocAppendix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12840
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton btnDelTemp 
      Caption         =   "Replace"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9480
      OleObjectBlob   =   "frmDocAppendix.frx":030A
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   300
      Left            =   11040
      TabIndex        =   2
      Top             =   240
      Width           =   1665
   End
   Begin VB.ComboBox cboTemplete 
      Height          =   300
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   4410
   End
   Begin VB.CommandButton btnSaveAs 
      Caption         =   "Replace"
      Height          =   375
      Left            =   9600
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   8295
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   14631
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Customer Pricing"
      TabPicture(0)   =   "frmDocAppendix.frx":2A0D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtRmk"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vendor Pricing"
      TabPicture(1)   =   "frmDocAppendix.frx":2A29
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnClear"
      Tab(1).Control(1)=   "btnItmDir"
      Tab(1).Control(2)=   "txtItmDir"
      Tab(1).Control(3)=   "imgCover"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   -68880
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton btnItmDir 
         Caption         =   "..."
         Height          =   315
         Left            =   -69360
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtItmDir 
         Height          =   300
         Left            =   -74760
         TabIndex        =   8
         Top             =   120
         Width           =   5355
      End
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
         Height          =   7695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   3
         Top             =   120
         Width           =   12075
      End
      Begin VB.Image imgCover 
         BorderStyle     =   1  '單線固定
         Height          =   7215
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   11775
      End
   End
   Begin MSComDlg.CommonDialog cdlgDir 
      Left            =   10560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDspDocNo 
      BorderStyle     =   1  '單線固定
      Caption         =   "NEW KEY:"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblDocNo 
      Caption         =   "NEW KEY:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDocAppendix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private wsFormID As String
Private waScrItm As New XArrayDB
Private wcCombo As Control

'variable for new property
Private msDocID    As String
Private msRmkID    As String
Private msRmkType  As String

Private wsRmkPrint  As String
Private Const wsPhotoPath = "..\Photo\"

Private wbExit    As Boolean
Private wsOldRmk    As String


Property Get DocID() As String

   DocID = msDocID
   
End Property

Property Let DocID(ByVal NewDocID As String)

   msDocID = NewDocID

End Property


Property Get RmkID() As String

   RmkID = msRmkID
   
End Property

Property Let RmkID(ByVal NewRmkID As String)

   msRmkID = NewRmkID

End Property
Property Let RmkType(ByVal NewRmkType As String)

   msRmkType = NewRmkType

End Property







Private Sub btnClear_Click()
    Call Clear_Cover
    Me.txtItmDir.Text = ""
    
End Sub

Private Sub btnDelTemp_Click()
    Call cmdSaveAs(1)
End Sub

Private Sub btnItmDir_Click()
    Dim wsFilePath As String
    
    
    cdlgDir.InitDir = wsPhotoPath
    cdlgDir.FileName = ""
    cdlgDir.ShowOpen
    wsFilePath = cdlgDir.FileName
    
    If Trim(wsFilePath) = "" Then Exit Sub
    
    If Chk_Load_Cover(wsFilePath) Then
        txtItmDir = wsFilePath
        'tabDetailInfo.Tab = 1
        'txtR
    End If
    
End Sub

Private Sub btnSaveAs_Click()
    Call cmdSaveAs(0)
End Sub

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
 Set frmDocAppendix = Nothing
    
End Sub





Private Sub tabDetailInfo_Click(PreviousTab As Integer)

    If tabDetailInfo.Tab = 0 Then
        
  '      If txtRmk.Enabled Then
  '          txtRmk.SetFocus
  '      End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
    
       ' If Me.tblCusItem.Enabled Then
       '     tblCusItem.SetFocus
    End If


End Sub



Private Sub tblCommon_DblClick()
    
    wcCombo.Text = tblCommon.Columns(0).Text
    tblCommon.Visible = False
    wcCombo.SetFocus
    SendKeys "{Enter}"
    
End Sub

Private Sub tblCommon_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        KeyCode = vbDefault
        tblCommon.Visible = False
        wcCombo.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = vbDefault
        wcCombo.Text = tblCommon.Columns(0).Text
        tblCommon.Visible = False
        wcCombo.SetFocus
        SendKeys "{Enter}"
    End If
        
End Sub


Private Sub tblCommon_LostFocus()
    
On Error GoTo tblCommon_LostFocus_Err
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If
    
Exit Sub
tblCommon_LostFocus_Err:
    Set wcCombo = Nothing

    
End Sub


Private Sub txtItmDir_GotFocus()
    FocusMe txtItmDir
End Sub

Private Sub txtItmDir_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmDir, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        If Trim(txtItmDir) = "" Then
            Clear_Cover
            btnItmDir.SetFocus
        Else
            If Chk_Load_Cover(txtItmDir) Then
                btnItmDir.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtItmDir_LostFocus()
    FocusMe txtItmDir, True
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
    wsFormID = "DOCAPP"
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    btnSaveAs.Caption = Get_Caption(waScrItm, "SAVEAS")
    btnDelTemp.Caption = Get_Caption(waScrItm, "DELTEMP")
    
    
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    
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
    
    tabDetailInfo.Tab = 0
    
    Call LoadRecord
    
    
End Sub

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    
    Select Case msRmkType
            Case "SO"
            wsSQL = "SELECT DAID, DARemark, DAName, DAPATH "
            wsSQL = wsSQL + "FROM mstDocAppendix "
            wsSQL = wsSQL + "WHERE DADOCID = " & msDocID
            
            lblDspDocNo.Caption = Get_TableInfo("SOASOHD", "SOHDDOCID =" & msDocID, "SOHDDOCNO")
            
            Case "SN"
            wsSQL = "SELECT DAID, DARemark, DAName, DAPATH "
            wsSQL = wsSQL + "FROM mstDocAppendix "
            wsSQL = wsSQL + "WHERE DADOCID = " & msDocID
            
            lblDspDocNo.Caption = Get_TableInfo("SOASNHD", "SNHDDOCID =" & msDocID, "SNHDDOCNO")
            
            
            Case Else
            
            wsSQL = "SELECT DAID, ARemark, DAName "
            wsSQL = wsSQL + "FROM mstDocAppendix "
            wsSQL = wsSQL + "WHERE DAID = " & msRmkID
            
    End Select
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
     '  Me.txtSaveAs = lblDspDocNo.Caption
        LoadRecord = False
    Else
      '  Me.txtSaveAs = ReadRs(rsRcd, "DAName")
        msRmkID = ReadRs(rsRcd, "DAID")
        Me.txtRmk = ReadRs(rsRcd, "DARemark")
        Me.txtItmDir = ReadRs(rsRcd, "DAPATH")
        wsOldRmk = txtRmk.Text
        
'        Call LoadItemImg(msRmkID)
        
        LoadRecord = True
    End If
    rsRcd.Close
    Set rsRcd = Nothing
    
    Call Clear_Cover
    
    If Trim(txtItmDir.Text) <> "" Then
        Call Load_Cover(Trim(txtItmDir.Text))
    End If
    
End Function

Public Function LoadTemplete() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    
    wsSQL = "SELECT DAID, DANAME, DARemark, DAPATH "
    wsSQL = wsSQL + "FROM mstDocAppendix "
    wsSQL = wsSQL + "WHERE DANAME = '" & Set_Quote(cboTemplete.Text) & "'"
            
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadTemplete = False
    Else
    '    Me.txtSaveAs = ReadRs(rsRcd, "DAName")
    
        msRmkID = ReadRs(rsRcd, "DAID")
        Me.txtRmk = ReadRs(rsRcd, "DARemark")
        Me.txtItmDir.Text = ReadRs(rsRcd, "DAPATH")
        wsOldRmk = txtRmk.Text
        
       ' Call LoadItemImg(msRmkID)
        LoadTemplete = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
        
    Call Clear_Cover
    
    If Trim(txtItmDir.Text) <> "" Then
        Call Load_Cover(Trim(txtItmDir.Text))
    End If
    
End Function


Private Function cmdSave() As Boolean
    
    Dim wsGenDte As String
    Dim iDel As Integer
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If Trim(txtRmk.Text) = "" Then
            iDel = 1
    Else
            iDel = 0
    End If
    
    If wsOldRmk = txtRmk.Text Then
            MousePointer = vbDefault
            Exit Function
    End If
    
    Call cmdReplace
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_DA001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, msDocID)
    Call SetSPPara(adcmdSave, 2, iDel)
    Call SetSPPara(adcmdSave, 3, msRmkType)
    Call SetSPPara(adcmdSave, 4, txtSaveAs.Text)
    Call SetSPPara(adcmdSave, 5, Trim(txtRmk.Text))
    Call SetSPPara(adcmdSave, 6, wsRmkPrint)
    Call SetSPPara(adcmdSave, 7, txtItmDir.Text)
    Call SetSPPara(adcmdSave, 8, gsUserID)
    Call SetSPPara(adcmdSave, 9, wsGenDte)
    adcmdSave.Execute
    msRmkID = GetSPPara(adcmdSave, 10)
    
    cnCon.CommitTrans
    
    If Trim(msRmkID) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - DOCAPPENDIX!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
    
       ' Call InsItemImg(txtItmDir.Text, msRmkID)
        
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


Private Function cmdSaveAs(iDel As Integer) As Boolean
    Dim wsGenDte As String

    Dim adcmdSaveAs As New ADODB.Command
    
    On Error GoTo cmdSaveAs_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    
    If iDel = 0 Then
    
    If Trim(txtSaveAs.Text) = "" Then
            gsMsg = "範本名稱沒有輸入!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
    
    Else
    
    If Trim(cboTemplete.Text) = "" Then
            gsMsg = "範本名稱沒有輸入!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
    
    
    End If
    
    If Trim(txtRmk.Text) = "" Then
            gsMsg = "內容沒有輸入!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
    
    
    
    cnCon.BeginTrans
    Set adcmdSaveAs.ActiveConnection = cnCon
        
    adcmdSaveAs.CommandText = "USP_DA001"
    adcmdSaveAs.CommandType = adCmdStoredProc
    adcmdSaveAs.Parameters.Refresh
      
    Call SetSPPara(adcmdSaveAs, 1, "")
    Call SetSPPara(adcmdSaveAs, 2, iDel)
    Call SetSPPara(adcmdSaveAs, 3, "SM")
    Call SetSPPara(adcmdSaveAs, 4, IIf(iDel = 1, cboTemplete.Text, txtSaveAs.Text))
    Call SetSPPara(adcmdSaveAs, 5, txtRmk.Text)
    Call SetSPPara(adcmdSaveAs, 6, wsRmkPrint)
    Call SetSPPara(adcmdSaveAs, 7, txtItmDir.Text)
    Call SetSPPara(adcmdSaveAs, 8, gsUserID)
    Call SetSPPara(adcmdSaveAs, 9, wsGenDte)
    adcmdSaveAs.Execute
    msRmkID = GetSPPara(adcmdSaveAs, 10)
    
    cnCon.CommitTrans
    
    If Trim(msRmkID) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - DOCAPPENDIX!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
         
      '   Call InsItemImg(txtItmDir.Text, msRmkID)
         
        gsMsg = "已成功儲存!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    End If
    
    
    Set adcmdSaveAs = Nothing
    cmdSaveAs = True
    
    MousePointer = vbDefault
    
    Exit Function
    
cmdSaveAs_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSaveAs = Nothing
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    
            Case vbKeyPageDown
            KeyCode = 0
            If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab + 1
                Exit Sub
            End If
        Case vbKeyPageUp
            KeyCode = 0
            If tabDetailInfo.Tab > 0 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab - 1
                Exit Sub
            End If
            
            
        Case vbKeyEscape
            wbExit = True
            Unload Me
    End Select
End Sub

Private Sub cmdReplace()

    Dim inText As String
    Dim replaceTxt As String
    Dim Totxt As String
    
    
    Me.MousePointer = vbHourglass
    inText = txtRmk.Text
    
    
        Select Case msRmkType
            Case "SO"
    
            replaceTxt = "@NETAMT"
            Totxt = Get_TableInfo("SOASOHD", "SOHDDOCID = " & msDocID, "SOHDNETAMT")
            Totxt = Format(Totxt, gsAmtFmt)
            inText = Replace_Str(inText, replaceTxt, Totxt)
    
            replaceTxt = "@DOCNO"
            Totxt = Get_TableInfo("SOASOHD", "SOHDDOCID = " & msDocID, "SOHDDOCNO")
            inText = Replace_Str(inText, replaceTxt, Totxt)
            
            replaceTxt = "@CURR"
            Totxt = Get_TableInfo("SOASOHD", "SOHDDOCID = " & msDocID, "SOHDCURR")
            inText = Replace_Str(inText, replaceTxt, Totxt)
            
            Case "SN"
    
            replaceTxt = "@NETAMT"
            Totxt = Get_TableInfo("SOASNHD", "SNHDDOCID = " & msDocID, "SNHDNETAMT")
            Totxt = Format(Totxt, gsAmtFmt)
            inText = Replace_Str(inText, replaceTxt, Totxt)
    
            replaceTxt = "@DOCNO"
            Totxt = Get_TableInfo("SOASNHD", "SNHDDOCID = " & msDocID, "SNHDDOCNO")
            inText = Replace_Str(inText, replaceTxt, Totxt)
            
            replaceTxt = "@CURR"
            Totxt = Get_TableInfo("SOASNHD", "SNHDDOCID = " & msDocID, "SNHDCURR")
            inText = Replace_Str(inText, replaceTxt, Totxt)
            
    End Select
    
    
    
   ' txtRmk.Text = inText
    wsRmkPrint = inText
    
    Me.MousePointer = vbNormal
    
    End Sub
    
    
Private Sub cboTemplete_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboTemplete
    
    wsSQL = "SELECT DANAME, Cast(DARemark as char(30)) "
    wsSQL = wsSQL & " FROM mstDocAppendix "
    wsSQL = wsSQL & " WHERE DAType  = 'SM' "
    wsSQL = wsSQL & " AND DAStatus = '1' "
        
    Call Ini_Combo(2, wsSQL, cboTemplete.Left, cboTemplete.Top + cboTemplete.Height, tblCommon, wsFormID, "TBLTEMP", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboTemplete_GotFocus()
    FocusMe cboTemplete
    Set wcCombo = cboTemplete
End Sub

Private Sub cboTemplete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboTemplete.Text) <> "" Then
             Call LoadTemplete
        End If
        
    End If
End Sub

Private Sub cboTemplete_LostFocus()
    FocusMe cboTemplete, True
End Sub

Private Sub txtSaveAs_GotFocus()
    FocusMe txtSaveAs
End Sub




Private Sub txtSaveAs_KeyPress(KeyAscii As Integer)
Call chk_InpLen(txtSaveAs, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        btnSaveAs.SetFocus
        
    End If
    
End Sub


Private Sub txtSaveAs_LostFocus()
    FocusMe txtSaveAs, True
End Sub



Private Function Clear_Cover() As Boolean
On Error GoTo Clear_Cover_Err

    Clear_Cover = False
    imgCover.Picture = LoadPicture()
    Clear_Cover = True
    Exit Function
    
Clear_Cover_Err:
    Clear_Cover = False
End Function


Private Function Chk_Load_Cover(inPath As String) As Boolean
    Chk_Load_Cover = False
    
    If Load_Cover(inPath) = False Then
        gsMsg = "封面圖象不存在或錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 1
        txtItmDir.SetFocus
        Exit Function
    End If
    
    Chk_Load_Cover = True
End Function

Private Function Load_Cover(inPath As String) As Boolean
On Error GoTo Load_Cover_Err

    Load_Cover = False
    imgCover.Picture = LoadPicture(inPath)
    Load_Cover = True
    Exit Function
    
Load_Cover_Err:
    Load_Cover = False
End Function


Private Function InsItemImg(ByVal inFileName As String, ByVal InDAId As Long) As Integer

On Error GoTo Err_Hanlder
     Dim rs As New ADODB.Recordset
     Dim stm As ADODB.Stream
     Dim wsSQL As String
     
     
     If Trim(inFileName) = "" Then
     Exit Function
     End If

     Set stm = New ADODB.Stream

    wsSQL = "Select * from mstDocAppendix Where DAID = " & InDAId

    rs.Open wsSQL, cnCon, adOpenKeyset, adLockOptimistic
        
  '   rs.Open "Select * from MstItemImg Where IMItemID = " & InItmID, m_dbh.GetConnection, adOpenKeyset, adLockOptimistic
     'Read the binary files from disk.
     stm.Type = adTypeBinary
     stm.Open
     stm.LoadFromFile inFileName
     
     If rs.RecordCount = 0 Then
     'rs.AddNew
     'rs!IMItemID = InItmID
     'rs!IMpath = inFileName
     'rs!IMImg = stm.Read
     'Insert the binary object into the table.
     'rs.Update
     InsItemImg = 0
     Else
     rs!DaPath = inFileName
     rs!DAImg = stm.Read
     'Insert the binary object into the table.
     rs.UPDATE
     
     End If
     
     

     rs.Close
     stm.Close
     
     InsItemImg = 1

     Set rs = Nothing
     Set stm = Nothing
     
     
     'Call m_dbh.RunSQL("Update MstItem Set ItmPackingSize = '" & Set_Key(inFileName) & "' Where ItmID = " & To_Num(InItmID), Array())
     
    Exit Function
    
Err_Hanlder:

    InsItemImg = 0
    MsgBox Err.Description & " -> Insert Photo Failed!"
     rs.Close
     stm.Close
     Set rs = Nothing
     Set stm = Nothing

 End Function
 
 
 
 Private Sub LoadItemImg(ByVal InDAId As Long)

On Error GoTo Err_Hanlder
     Dim rs As New ADODB.Recordset
     Dim wsSQL As String
     Dim DataFile As Integer, Fl As Long, Chunks As Integer
     Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
     Dim ChunkSize As Integer, conChunkSize As Integer
     
     ChunkSize = 16384
     conChunkSize = 100
     
     Call Clear_Cover
    
    wsSQL = "Select * from mstDocAppendix Where DAID = " & InDAId

    rs.Open wsSQL, cnCon, adOpenForwardOnly, adLockReadOnly
     If rs.RecordCount <> 0 Then
     
    DataFile = 1
    Open "pictemp" For Binary Access Write As DataFile
        Fl = rs!DAImg.ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = rs!DAImg.GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = rs!DAImg.GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = "pictemp"
    imgCover.Picture = LoadPicture(FileName)
     

     End If
     

     rs.Close
     Set rs = Nothing
     
    Exit Sub
    
Err_Hanlder:


    MsgBox Err.Description & " -> Insert Photo Failed!"
     rs.Close
     Set rs = Nothing
     
 End Sub



