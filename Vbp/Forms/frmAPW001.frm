VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmAPW001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPW001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmAPW001.frx":0442
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboWorkNo 
      Height          =   300
      Left            =   9360
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5415
      Left            =   360
      OleObjectBlob   =   "frmAPW001.frx":2B45
      TabIndex        =   13
      Top             =   2640
      Width           =   11295
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   6015
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAPW001.frx":D29C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmAPW001.frx":D2B8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmAPW001.frx":D2D4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.ComboBox cboStaffNo 
      Height          =   300
      Left            =   9360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11775
      Begin VB.Frame fraSelect 
         Height          =   1485
         Left            =   7320
         TabIndex        =   7
         Top             =   120
         Width           =   4215
         Begin VB.Label lblWorkNo 
            Caption         =   "Customer Code From"
            Height          =   225
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1650
         End
         Begin VB.Label lblStaffNo 
            Caption         =   "Customer Code From"
            Height          =   225
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1650
         End
      End
      Begin VB.Label lblDspJobRef3 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label lblDspJobRef2 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lblDspJobRef1 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label lblJobRef 
         Caption         =   "CUSNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Width           =   1890
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   11400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":D2F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":DBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":E4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":E8F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":ED48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":F062
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":F4B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":F906
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":FC20
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":FF3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":1038C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":10C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":10F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":113E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":11700
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":11A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":11E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":1218C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":124AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":127CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":12C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPW001.frx":12F40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Can"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Finish"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmAPW001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
'Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Private wbErr As Boolean



Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer
Private wsMark As String
Private wsRTitle As String
Private wsDteTim As String


Private wlKey As Long
Private wlStaffID As Long
Private wlWorkID As Long
Private wsWhsCode As String

Private wlLastRow As Integer

Private Const tcConvert = "Convert"
Private Const tcCan = "Can"
Private Const tcFinish = "Finish"
Private Const tcPrint = "Print"
Private Const tcExport = "Export"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcAdd = "Add"

Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const SDOCLINE = 1
Private Const SDOCNO = 2
Private Const SITMCODE = 3
Private Const SITMNAME = 4
Private Const SITMTYPE = 5
Private Const SLOTNO = 6
Private Const SQTY = 7
Private Const SOUTQTY = 8
Private Const SREM = 9
Private Const SAPRFLG = 10
Private Const SREM2 = 11
Private Const SWHS2 = 12
Private Const SREM3 = 13
Private Const SWHS3 = 14
Private Const SREM4 = 15
Private Const SWHS4 = 16
Private Const SID = 17



Private Sub cboStaffNo_DropDown()
    Dim wsSQL As String
    
    
    
    wsSQL = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboStaffNo.SelLength > 0, "", Set_Quote(cboStaffNo.Text)) & "%' "
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "AND SaleType = 'W' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
        
    
    Call Ini_Combo(2, wsSQL, cboStaffNo.Left, cboStaffNo.Top + cboStaffNo.Height, tblCommon, wsFormID, "TBLSTAFFNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboStaffNo_GotFocus()
        FocusMe cboStaffNo
    Set wcCombo = cboStaffNo
End Sub

Private Sub cboStaffNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboStaffNo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboStaffNo = False Then Exit Sub
        
        cboWorkNo.SetFocus
        
    End If
End Sub


Private Sub cboStaffNo_LostFocus()
    FocusMe cboStaffNo, True
End Sub

Private Function chk_cboStaffNo() As Boolean
Dim wsName As String

 chk_cboStaffNo = False
    
 If Chk_Salesman(cboStaffNo.Text, wlStaffID, wsName) = False Then
        gsMsg = "Satff Not Exist!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboStaffNo.SetFocus
        Exit Function
  End If
  
 
    
  chk_cboStaffNo = True
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub

Private Sub cboDocNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    
    
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '1' "
    wsSQL = wsSQL & " AND SOHDCTLPRD BETWEEN '" & Str(Val(Left(gsSystemDate, 4)) - 1) + "01" & "' AND '" & Left(gsSystemDate, 4) + "12" & "'"
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    
    Call Ini_Combo(3, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoFr = False Then Exit Sub
        
        Call LoadRecord
        cboStaffNo.SetFocus
        
    End If
End Sub

Private Function chk_cboDocNoFr() As Boolean
Dim wsStatus As String
    chk_cboDocNoFr = False
    
 If Chk_TrnHdDocNo("SO", cboDocNoFr, wsStatus) = False Then
        gsMsg = "Job No Not Exist!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNoFr.SetFocus
        Exit Function
  Else
  
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNoFr.SetFocus
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNoFr.SetFocus
            Exit Function
        End If
        
  End If
  
  Get_RefDoc
    
  chk_cboDocNoFr = True
End Function


 
Private Sub Chk_Sel(inRow As Long)
    
    Dim wlCtr As Long
     
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If waResult(wlCtr, SSEL) = "-1" Then
                  waResult(wlCtr, SSEL) = "0"
                  Exit Sub
               End If
            End If
        Next

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF10
        If tbrProcess.Buttons(tcConvert).Enabled = False Then Exit Sub
           Call cmdPick(1)
        
        Case vbKeyF3
        If tbrProcess.Buttons(tcCan).Enabled = False Then Exit Sub
           Call cmdPick(2)
           
        Case vbKeyF2
        If tbrProcess.Buttons(tcAdd).Enabled = False Then Exit Sub
           Call cmdAddItem
           
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
             
        Case vbKeyF5
           Call cmdSelect(1)
           
        Case vbKeyF6
           Call cmdSelect(0)
           
                   
        Case vbKeyF7
            Call cmdRefresh
            
            
        Case vbKeyF9
        If tbrProcess.Buttons(tcPrint).Enabled = False Then Exit Sub
            Call cmdPrint
            
        
        
            
    End Select
End Sub









Private Sub tabDetailInfo_Click(PreviousTab As Integer)



Call cmdRefresh


End Sub



Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    Select Case tblDetail.Col
        
        Case SREM, SREM2, SREM3, SREM4
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
   
       
    End Select
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    
   If tbrProcess.Buttons(Button.Key).Enabled = False Then Exit Sub
        
    
    Select Case Button.Key
        Case tcConvert
            Call cmdPick(1)
            
        Case tcCan
            Call cmdPick(2)
                 
        Case tcAdd
            Call cmdAddItem
        
        Case tcPrint
            Call cmdPrint
        
        Case tcExport
            Call cmdExport
        
        Case tcCancel
        
           Call cmdCancel
            
        
        Case tcSAll
        
           Call cmdSelect(1)
        
        Case tcDAll
        
           Call cmdSelect(0)
           
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call cmdRefresh
            
            
    End Select
End Sub

Private Sub Form_Load()
    
    
  MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Grid
    Ini_Scr

    
   MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCancel()
    
    
  MousePointer = vbHourglass
  
    Ini_Scr
    
   MousePointer = vbDefault
    
    
End Sub



Private Sub cmdRefresh()
    
    
  MousePointer = vbHourglass
  
    Call Set_tbrProcess
    Call LoadRecord
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, SSEL, SID
    
    
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
   '         Case "ComboBox"
   '             MyControl.Clear
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

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    wiExit = False
    
    
    
    
    cboDocNoFr.Text = ""
    cboStaffNo.Text = ""
    cboWorkNo.Text = ""
        
    wsWhsCode = Get_TableInfo("SYSWSINFO", "WSID = '01'", "WSWHSCODE")


     Call Set_tbrProcess

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   
    cnCon.Execute "DELETE FROM RPTAPW001 WHERE RPTUSRID = '" & gsUserID & "' AND RPTDTETIM = '" & wsDteTim & "' "
    
    Set waScrItm = Nothing
 '   Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmAPW001 = Nothing
 
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
 '   wsFormID = "APW001"
    wsDteTim = Change_SQLDate(Now)
    
End Sub


Private Sub Set_tbrProcess()

Dim wiCtr As Integer

With tbrProcess
    
    Select Case tabDetailInfo.Tab
    Case 0
    
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcAdd).Enabled = True
    .Buttons(tcPrint).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcExport).Enabled = True
    
     Case 1
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = True
    .Buttons(tcAdd).Enabled = False
    .Buttons(tcPrint).Enabled = True
    .Buttons(tcFinish).Enabled = True
    .Buttons(tcExport).Enabled = True
     
     Case 2
    
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcAdd).Enabled = False
    .Buttons(tcPrint).Enabled = True
    .Buttons(tcFinish).Enabled = True
    .Buttons(tcExport).Enabled = True
    
    End Select
    
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    

    
End With


With tblDetail

Select Case tabDetailInfo.Tab
    Case 0

    For wiCtr = SSEL To SID
    Select Case wiCtr
                Case SDOCNO
                    .Columns(wiCtr).Width = 700
                Case SITMTYPE
                    .Columns(wiCtr).Visible = False
                Case SREM
                    .Columns(wiCtr).Caption = Get_Caption(waScrItm, "SREM1")
                Case SAPRFLG
                    .Columns(wiCtr).Caption = Get_TableInfo("sysCmpWhs", "CWhsID = 1", IIf(gsLangID = "1", "CWhsDesc1", "CWhsChinDesc1"))
                Case SREM2
                    .Columns(wiCtr).Visible = True
                Case SWHS2
                    .Columns(wiCtr).Visible = True
                Case SREM3
                    .Columns(wiCtr).Visible = True
                Case SWHS3
                    .Columns(wiCtr).Visible = True
                Case SREM4
                    .Columns(wiCtr).Visible = True
                Case SWHS4
                    .Columns(wiCtr).Visible = True
                Case SLOTNO
                    .Columns(wiCtr).Visible = True
             
    End Select
    Next wiCtr
 
 Case 1, 2
    For wiCtr = SSEL To SID
    Select Case wiCtr
                Case SDOCNO
                    .Columns(wiCtr).Width = 1200
                Case SITMTYPE
                    .Columns(wiCtr).Visible = True
                Case SREM
                    .Columns(wiCtr).Caption = Get_Caption(waScrItm, "SREM")
                Case SAPRFLG
                    .Columns(wiCtr).Caption = Get_Caption(waScrItm, "SAPRFLG")
                Case SREM2
                    .Columns(wiCtr).Visible = False
                Case SWHS2
                    .Columns(wiCtr).Visible = False
                Case SREM3
                    .Columns(wiCtr).Visible = False
                Case SWHS3
                    .Columns(wiCtr).Visible = False
                Case SREM4
                    .Columns(wiCtr).Visible = False
                Case SWHS4
                    .Columns(wiCtr).Visible = False
                Case SLOTNO
                    .Columns(wiCtr).Visible = False
                
                
    End Select
    Next wiCtr
 

 
End Select
     

End With
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
  '  Call Get_Scr_Item("TOOLTIP_A", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblStaffNo.Caption = Get_Caption(waScrItm, "STAFFNO")
    lblWorkNo.Caption = Get_Caption(waScrItm, "WORKNO")
    
    lblJobRef.Caption = Get_Caption(waScrItm, "JOBREF")
    
    wsRTitle = Get_Caption(waScrItm, "TITLE")
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SLOTNO).Caption = Get_Caption(waScrItm, "SLOTNO")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SOUTQTY).Caption = Get_Caption(waScrItm, "SOUTQTY")
        .Columns(SREM).Caption = Get_Caption(waScrItm, "SREM")
        .Columns(SAPRFLG).Caption = Get_Caption(waScrItm, "SAPRFLG")
        .Columns(SREM2).Caption = Get_Caption(waScrItm, "SREM1")
        .Columns(SWHS2).Caption = Get_TableInfo("sysCmpWhs", "CWhsID = 1", IIf(gsLangID = "1", "CWhsDesc2", "CWhsChinDesc2"))
        .Columns(SREM3).Caption = Get_Caption(waScrItm, "SREM1")
        .Columns(SWHS3).Caption = Get_TableInfo("sysCmpWhs", "CWhsID = 1", IIf(gsLangID = "1", "CWhsDesc3", "CWhsChinDesc3"))
        .Columns(SREM4).Caption = Get_Caption(waScrItm, "SREM1")
        .Columns(SWHS4).Caption = Get_TableInfo("sysCmpWhs", "CWhsID = 1", IIf(gsLangID = "1", "CWhsDesc4", "CWhsChinDesc4"))
        
        
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "OPT1")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "OPT2")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "OPT3")
    
    
    
    With tbrProcess
    .Buttons(tcConvert).ToolTipText = Get_Caption(waScrItm, tcConvert) & "(F10)"
    .Buttons(tcCan).ToolTipText = Get_Caption(waScrItm, tcCan) & "(F3)"
    .Buttons(tcPrint).ToolTipText = Get_Caption(waScrItm, tcPrint) & "(F9)"
    
    
    .Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F7)"
    .Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F11)"
    .Buttons(tcAdd).ToolTipText = Get_Caption(waScrItm, tcAdd) & "(F2)"
    
    .Buttons(tcSAll).ToolTipText = Get_Caption(waScrItm, tcSAll) & "(F5)"
    .Buttons(tcDAll).ToolTipText = Get_Caption(waScrItm, tcDAll) & "(F6)"
    .Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
   End With

End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
    End With

   If ColIndex = SSEL Then
   
 '   tblDetail.ReBind
 '   tblDetail.Bookmark = 0
         
   End If
   
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsLotNo As String


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case SSEL
            
           '   If .Columns(ColIndex).Text = "-1" Then
           '       Call Chk_Sel(.Row + To_Value(.FirstRow))
           '    End If
           
           Case SLOTNO
           
           
           If tabDetailInfo.Tab = 0 Then
           
                If Chk_LotEnabled("CF-HK") = True Then
                     .Columns(SAPRFLG).Text = Get_LocBalbyCode(.Columns(SITMCODE).Text, "CF-HK", Trim(.Columns(ColIndex).Text))
                End If
                
                If Chk_LotEnabled("CF-HZ") = True Then
                     .Columns(SWHS2).Text = Get_LocBalbyCode(.Columns(SITMCODE).Text, "CF-HZ", Trim(.Columns(ColIndex).Text))
                End If
                
                If Chk_LotEnabled("CF-GG") = True Then
                     .Columns(SWHS3).Text = Get_LocBalbyCode(.Columns(SITMCODE).Text, "CF-GG", Trim(.Columns(ColIndex).Text))
                End If
                
                If Chk_LotEnabled("WHS001") = True Then
                     .Columns(SWHS4).Text = Get_LocBalbyCode(.Columns(SITMCODE).Text, "WHS001", Trim(.Columns(ColIndex).Text))
                End If
                
                
                If Chk_grdLotNo("CF-HK", .Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                            
                
            End If
                
                           
                
            Case SREM
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case SREM2
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case SREM3
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case SREM4
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                
            End Select
    End With
    
    Exit Sub
    
Tbl_BeforeColUpdate_Err:
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True
    Exit Sub

tblDetail_BeforeColUpdate_Err:
    
    MsgBox "Check tblDeiail BeforeColUpdate!"
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True
End Sub



Private Sub tblDetail_ButtonClick(ByVal ColIndex As Integer)

    
    Dim wsSQL As String
    Dim wiTop As Long
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err
    
    With tblDetail
        Select Case ColIndex
        Case SLOTNO
        
            wsSQL = "SELECT ILLOCCODE, ILSOHQTY "
            wsSQL = wsSQL & " FROM ICLOCBAL, MSTITEM "
            wsSQL = wsSQL & " WHERE ILLOCCODE LIKE '%" & Set_Quote(.Columns(SLOTNO).Text) & "%' "
            wsSQL = wsSQL & " AND ILITEMID = ITMID "
            wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(.Columns(SITMCODE).Text) & "'"
            wsSQL = wsSQL & " AND ILWHSCODE = '" & wsWhsCode & "'"
    
             Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLLOCCODE", Me.Width, Me.Height)
             tblCommon.Visible = True
             tblCommon.SetFocus
             Set wcCombo = tblDetail
             
        
           End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDetail ButtonClick!"
 
    
 
End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
        Case vbKeyF4        ' CALL COMBO BOX
            KeyCode = vbDefault
            Call tblDetail_ButtonClick(.Col)
            
        Case vbKeyReturn
            Select Case .Col
            Case SREM, SWHS4
                 KeyCode = vbKeyDown
                 .Col = SSEL
            Case SSEL
                 KeyCode = vbDefault
                 .Col = SREM
            Case Else
                 KeyCode = vbDefault
                 .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SSEL Then
                .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SREM, SWHS4
                    KeyCode = vbKeyDown
                    .Col = SSEL
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
                
            End Select
        
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub






Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
    lblDspItmDesc.Caption = .Columns(.Col).Text
    End With
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
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


Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = dbgDisabled
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SSEL To SID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SSEL
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).Locked = False
                Case SDOCLINE
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Width = 500
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 1200
                Case SITMCODE
                    .Columns(wiCtr).Width = 3000
                   .Columns(wiCtr).DataWidth = 30
                Case SITMNAME
                   .Columns(wiCtr).Width = 2500
                   .Columns(wiCtr).DataWidth = 60
                Case SITMTYPE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SLOTNO
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).HeadForeColor = vbRed
                    .Columns(wiCtr).Button = True
                Case SQTY
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SOUTQTY
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SREM
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).HeadForeColor = vbRed
                Case SAPRFLG
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SREM2
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).HeadForeColor = vbRed
                Case SWHS2
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SREM3
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).HeadForeColor = vbRed
                Case SWHS3
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SREM4
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).HeadForeColor = vbRed
                Case SWHS4
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
 
                Case SID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 15
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
    
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wdCreLmt As Double
    Dim wdCreLft As Double
    Dim wsStatus As String
    Dim wdQty As Double
    Dim wdSpQty As Double
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    Call cmdSave
       
    
    wsSQL = "SELECT RPTDOCLINE, RPTDOCID, RPTDOCNO, RPTSID, RPTITMID, RPTITMCODE, RPTITMNAME, RPTITMTYPE, "
    wsSQL = wsSQL & " RPTQTY, RPTTRQTY, RPTREMQTY, RPTSOH, RPTAPRFLG, "
    wsSQL = wsSQL & " RPTREM2, RPTSOH2, RPTREM3, RPTSOH3, RPTREM4, RPTSOH4"
    wsSQL = wsSQL & " FROM RPTAPW001 "
    wsSQL = wsSQL & " WHERE RPTUSRID = '" & gsUserID & "' "
    wsSQL = wsSQL & " AND RPTDTETIM = '" & wsDteTim & "' "
    If tabDetailInfo.Tab = 0 Then
    wsSQL = wsSQL & " ORDER BY RPTDOCLINE, CONVERT(INTEGER, RPTDOCNO)  "
    Else
    wsSQL = wsSQL & " ORDER BY RPTDOCLINE, RPTDOCNO  "
    End If
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SSEL, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
     
    With waResult
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    
     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCLINE) = Format(ReadRs(rsRcd, "RPTDOCLINE"), "000")
        waResult(.UpperBound(1), SDOCNO) = IIf(tabDetailInfo.Tab = 0, Format(ReadRs(rsRcd, "RPTDOCNO"), "000"), ReadRs(rsRcd, "RPTDOCNO"))
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "RPTITMCODE")
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "RPTITMNAME")
        waResult(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "RPTITMTYPE")
        waResult(.UpperBound(1), SLOTNO) = ReadRs(rsRcd, "RPTLOTNO")
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "RPTQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SOUTQTY) = Format(To_Value(ReadRs(rsRcd, "RPTTRQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SREM) = Format(To_Value(ReadRs(rsRcd, "RPTREMQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SAPRFLG) = IIf(tabDetailInfo.Tab = 0, Format(To_Value(ReadRs(rsRcd, "RPTSOH")), gsQtyFmt), ReadRs(rsRcd, "RPTAPRFLG"))
        waResult(.UpperBound(1), SREM2) = Format(To_Value(ReadRs(rsRcd, "RPTREM2")), gsQtyFmt)
        waResult(.UpperBound(1), SWHS2) = Format(To_Value(ReadRs(rsRcd, "RPTSOH2")), gsQtyFmt)
        waResult(.UpperBound(1), SREM3) = Format(To_Value(ReadRs(rsRcd, "RPTREM3")), gsQtyFmt)
        waResult(.UpperBound(1), SWHS3) = Format(To_Value(ReadRs(rsRcd, "RPTSOH3")), gsQtyFmt)
        waResult(.UpperBound(1), SREM4) = Format(To_Value(ReadRs(rsRcd, "RPTREM4")), gsQtyFmt)
        waResult(.UpperBound(1), SWHS4) = Format(To_Value(ReadRs(rsRcd, "RPTSOH4")), gsQtyFmt)
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "RPTSID")
        
     
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    rsRcd.Close
    Set rsRcd = Nothing
 
 
    
    LoadRecord = True
    Me.MousePointer = vbNormal
    
End Function


Private Function Chk_GrdRow(ByVal LastRow As Long) As Boolean

    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    
    Chk_GrdRow = False
    
    On Error GoTo Chk_GrdRow_Err
    
    With tblDetail
        
        If To_Value(LastRow) > waResult.UpperBound(1) Then
           Chk_GrdRow = True
           Exit Function
        End If
        
        
        
        If wiActFlg = 1 Then
        If tabDetailInfo.Tab = 0 Then
        
        If To_Value(waResult(LastRow, SREM)) = 0 And _
           To_Value(waResult(LastRow, SREM2)) = 0 And _
           To_Value(waResult(LastRow, SREM3)) = 0 And _
           To_Value(waResult(LastRow, SREM4)) = 0 Then
                .Col = SITMCODE
               gsMsg = "數量必需大於零"
               MsgBox gsMsg, vbOKOnly, gsTitle
               Exit Function
        End If
        
        
        If To_Value(waResult(LastRow, SREM)) > 0 And To_Value(waResult(LastRow, SAPRFLG)) < To_Value(waResult(LastRow, SREM)) Then
            .Col = SAPRFLG
             gsMsg = "沒有足夠物料(HK)!不能轉移"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        If To_Value(waResult(LastRow, SREM2)) > 0 And To_Value(waResult(LastRow, SWHS2)) < To_Value(waResult(LastRow, SREM2)) Then
            .Col = SWHS2
             gsMsg = "沒有足夠物料(HZ)!不能轉移"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        If To_Value(waResult(LastRow, SREM3)) > 0 And To_Value(waResult(LastRow, SWHS3)) < To_Value(waResult(LastRow, SREM3)) Then
            .Col = SWHS3
             gsMsg = "沒有足夠物料(3)!不能轉移"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        If To_Value(waResult(LastRow, SREM4)) > 0 And To_Value(waResult(LastRow, SWHS4)) < To_Value(waResult(LastRow, SREM4)) Then
            .Col = SWHS4
             gsMsg = "沒有足夠物料(4)!不能轉移"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        
        Else
        
        If Chk_grdQty(waResult(LastRow, SREM)) = False Then
                .Col = SREM
               Exit Function
        End If
        
        If waResult(LastRow, SAPRFLG) = "N" Then
            .Col = SAPRFLG
             gsMsg = "沒有批核!不能轉移"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        End If
        
        Else
        
        If Chk_grdQty(waResult(LastRow, SREM)) = False Then
                .Col = SREM
               Exit Function
        End If
        
        If waResult(LastRow, SAPRFLG) = "N" And tabDetailInfo.Tab = 2 Then
            .Col = SAPRFLG
             gsMsg = "已批核!不能刪除"
             MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        
        End If
        
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function


Private Function Chk_grdQty(inCode As String) As Boolean
    
    Chk_grdQty = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdQty = False
        Exit Function
    End If

    If To_Value(inCode) < 0 Then
        gsMsg = "數量必需大於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdQty = False
        Exit Function
    End If
    
    
    
    
End Function


Private Function InputValidation() As Boolean
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    wlLastRow = 0
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SSEL)) = "-1" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
                wlLastRow = wlLastRow + 1
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
        tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    If chk_cboDocNoFr = False Then Exit Function
    
    If wiActFlg = 1 Then
    
    
    If chk_cboStaffNo = False Then
        Exit Function
    End If
    
    If tabDetailInfo.Tab = 1 Then
    
    If chk_cboWorkNo = False Then
        Exit Function
    End If
    
    End If
    
    End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function





Private Sub cmdSelect(ByVal wiSelect As Integer)
    Dim wiCtr As Long
    
    Me.MousePointer = vbHourglass
    
    
     
    With waResult
    For wiCtr = 0 To .UpperBound(1)
        waResult(wiCtr, SSEL) = IIf(wiSelect = 1, "-1", "0")
    Next wiCtr
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    Me.MousePointer = vbNormal
    
End Sub


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property




Private Function Get_RefDoc() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Get_RefDoc = False
    
        wsSQL = "SELECT SOHDDOCID, SOHDSHIPFROM, SOHDSHIPTO, SOHDSHIPVIA "
        wsSQL = wsSQL & "FROM  soaSOHD "
        wsSQL = wsSQL & "WHERE SOHDDOCNO = '" & Set_Quote(cboDocNoFr) & "' "
        
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlKey = To_Value(ReadRs(rsRcd, "SOHDDOCID"))
    lblDspJobRef1 = ReadRs(rsRcd, "SOHDSHIPFROM")
    lblDspJobRef2 = ReadRs(rsRcd, "SOHDSHIPTO")
    lblDspJobRef3 = ReadRs(rsRcd, "SOHDSHIPVIA")
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Get_RefDoc = True
    
End Function

Private Sub cmdPick(ByVal inActFlg As Integer)

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlLineNo As Long
    Dim wlHDID As Long
    Dim wsTrnCd As String
    Dim wsWhsNo As String
     
    On Error GoTo cmdPick_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    wiActFlg = inActFlg
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    
    If wiActFlg = 1 Then
    gsMsg = "你是否確認要轉換工序?"
    Select Case tabDetailInfo.Tab
    Case 0
    wsTrnCd = "SP"
    Case 1
    wsTrnCd = "SW"
    End Select
    Else
    gsMsg = "你是否確認刪除物料?"
    Select Case tabDetailInfo.Tab
    Case 1
    wsTrnCd = "SP"
    Case 2
    wsTrnCd = "SW"
    End Select
    End If
    
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

 
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    wlLineNo = 1
    wlHDID = 0
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_APW001A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wiActFlg)
                Call SetSPPara(adcmdSave, 2, wsTrnCd)
                Call SetSPPara(adcmdSave, 3, wlKey)
                Call SetSPPara(adcmdSave, 4, wlHDID)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 6, wlLineNo)
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, SREM))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, SREM2))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, SREM3))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, SREM4))
                Call SetSPPara(adcmdSave, 11, wlStaffID)
                Call SetSPPara(adcmdSave, 12, wlWorkID)
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, SLOTNO))
                Call SetSPPara(adcmdSave, 14, wsFormID)
                Call SetSPPara(adcmdSave, 15, gsUserID)
                Call SetSPPara(adcmdSave, 16, wsGenDte)
                Call SetSPPara(adcmdSave, 17, IIf(wlLastRow = wlLineNo, "Y", "N"))
                
                adcmdSave.Execute
                wlHDID = GetSPPara(adcmdSave, 18)
                wsDocNo = GetSPPara(adcmdSave, 19)
                
                If wlHDID < 0 Then
                wsDocNo = waResult(wiCtr, SDOCNO)
                GoTo USP_APW001A_Err
                End If
                wlLineNo = wlLineNo + 1
            End If
        Next
    End If
    
     
    cnCon.CommitTrans
    
  
    gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
        

    
    Set adcmdSave = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    Exit Sub
        
    
USP_APW001A_Err:

    If wiActFlg = 1 And wsTrnCd = "SP" Then
    Select Case wlHDID
    Case -1
        wsWhsNo = "香港倉"
    Case -2
        wsWhsNo = "鶴山倉"
    Case -3
        wsWhsNo = "第三倉"
    Case -4
        wsWhsNo = "第四倉"
    End Select
    
    gsMsg = "物料" & wsDocNo & "在" & wsWhsNo & "不足夠!不能轉移"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    ElseIf wiActFlg = 2 And wsTrnCd = "SP" Then
    
    gsMsg = "物料在倉B不足夠!不能刪除"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    ElseIf wiActFlg = 2 And wsTrnCd = "SW" Then
    
    Else
    
    
    MsgBox Err.Description
    End If
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
    Exit Sub

cmdPick_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub

Private Sub cboWorkNo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT SaleCode, SaleName FROM mstsalesman WHERE SaleCode LIKE '%" & IIf(cboWorkNo.SelLength > 0, "", Set_Quote(cboWorkNo.Text)) & "%' "
    wsSQL = wsSQL & " AND SaleStatus <> '2' "
    wsSQL = wsSQL & "AND SaleType = 'W' "
    wsSQL = wsSQL & " ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboWorkNo.Left, cboWorkNo.Top + cboWorkNo.Height, tblCommon, wsFormID, "TBLWorkNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboWorkNo_GotFocus()
        FocusMe cboWorkNo
    Set wcCombo = cboWorkNo
End Sub

Private Sub cboWorkNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWorkNo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboWorkNo = False Then Exit Sub
        
        tblDetail.SetFocus
        
    End If
End Sub


Private Sub cboWorkNo_LostFocus()
    FocusMe cboWorkNo, True
End Sub

Private Function chk_cboWorkNo() As Boolean
Dim wsName As String

 chk_cboWorkNo = False
    
 If Chk_Salesman(cboWorkNo.Text, wlWorkID, wsName) = False Then
        gsMsg = "Worker Not Exist!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboWorkNo.SetFocus
        Exit Function
  End If
  
 
    
  chk_cboWorkNo = True
End Function


Private Sub cmdAddItem()
Dim wiCtr As Integer

 
  If Trim(cboDocNoFr) = "" Then
        gsMsg = "Job No Must Be Enter!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNoFr.SetFocus
        Exit Sub
 End If
 
 If Chk_SoReady(cboDocNoFr.Text) = True Then
    gsMsg = "文件已封存(Ready), 現在以唯讀模式開啟!請以密碼解封"
    MsgBox gsMsg, vbOKOnly, gsTitle
    cboDocNoFr.SetFocus
    Exit Sub
 End If

 frmAPW0011.KeyID = Get_TableInfo("SOASOHD", "SOHDDOCNO = '" & Set_Quote(cboDocNoFr.Text) & "'", "SOHDDOCID")
 frmAPW0011.Show vbModal
            
 If frmAPW0011.Result = True Then
      Call cmdRefresh
 End If
 
 Set frmAPW0011 = Nothing
            
      
End Sub


Private Sub cmdPrint()
    Dim wpDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    
    If Trim(cboDocNoFr) = "" Then
     gsMsg = "沒有選擇工程單!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    Exit Sub
    End If
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(1)
    wsSelection(1) = ""
    
    
    'Create Stored Procedure String
    wpDteTim = Change_SQLDate(Now)
    wsSQL = "EXEC usp_RPTAPW001 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wpDteTim) & "', "
    wsSQL = wsSQL & "" & tabDetailInfo.Tab & ", "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & gsLangID
        
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTAPW001"
    Else
    wsRptName = "RPTAPW001"
    End If
    
    If tabDetailInfo.Tab = 0 Then
    wsRptName = wsRptName + "A"
    End If
    
    NewfrmPrint.ReportID = "APW001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "APW001"
    NewfrmPrint.RptDteTim = wpDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdSave()
    Dim adcmdSave As New ADODB.Command

     
    On Error GoTo cmdSave_Err
    
    'MousePointer = vbHourglass
    
    
        
    If chk_cboDocNoFr = False Then
        cboDocNoFr.SetFocus
        Exit Sub
    End If
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    adcmdSave.CommandText = "USP_RPTAPW001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
     
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, wsDteTim)
    Call SetSPPara(adcmdSave, 3, tabDetailInfo.Tab)
    Call SetSPPara(adcmdSave, 4, cboDocNoFr)
    Call SetSPPara(adcmdSave, 5, gsLangID)
    
    adcmdSave.Execute
        
    cnCon.CommitTrans
    
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Set adcmdSave = Nothing
    
    
  '  MousePointer = vbDefault
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub


Private Sub cmdExport()

    Dim wsGenDte As String
    Dim wiCtr As Integer
    Dim wsTrnCode As String
    Dim wiMod As Integer
    Dim wsPath As String
    
    
     
    On Error GoTo cmdExport_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate

    If chk_cboDocNoFr = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
   
    gsMsg = "你是否確認要匯出文件？"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
    Select Case tabDetailInfo.Tab
    Case 0
        wsTrnCode = "SO"
    Case 1
        wsTrnCode = "SP"
    Case 2
        wsTrnCode = "SW"
    End Select
    
    

    If Trim(gsHHPath) <> "" Then
        wsPath = gsHHPath + "send\HHTORDER.TXT"
    Else
        wsPath = App.Path + "send\HHTORDER.TXT"
    End If
    
    
    wiMod = 1
    If ExportToHHFile(wsPath, wsTrnCode, wlKey, wiMod, "") = False Then
        gsMsg = cboDocNoFr.Text & " 匯出Error!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    
    
    Sleep (500)
    If SendToHH(wsPath) = True Then
  
    gsMsg = "匯出文件已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    End If
    
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    Exit Sub
    
cmdExport_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    
End Sub

Private Function Chk_grdLotNo(inWhs As String, inNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdLotNo = False
    
    If Chk_LotEnabled(inWhs) = False Then
        Chk_grdLotNo = True
        Exit Function
    End If
    
    
    If Chk_LotB(inWhs, inNo) = False Then
         gsMsg = "不能輸入 " & inNo & " 貨架!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
        
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入貨架!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdLotNo = True
    
End Function
