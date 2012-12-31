VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmMRP002 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMRP002.frx":0000
   LinkTopic       =   "Form1"
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
      OleObjectBlob   =   "frmMRP002.frx":0442
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5415
      Left            =   360
      OleObjectBlob   =   "frmMRP002.frx":2B45
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMRP002.frx":B018
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMRP002.frx":B034
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
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
         Begin VB.Label lblDspName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblStaffNo 
            Caption         =   "Customer Code From"
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   240
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
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":B050
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":B92A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":C204
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":C656
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":CAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":CDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":D214
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":D666
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":D980
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":DC9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":E0EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":E9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":ECF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":F144
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":F460
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":F77C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":FBD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":FEEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":1020C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP002.frx":1052C
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
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Pick"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Finish"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "frmMRP002"
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

Private wsMark As String

Private wlKey As Long
Private wlStaffID As Long
Private wlLastRow As Integer
Private wiActFlg As Integer


Private Const tcConvert = "Convert"
Private Const tcPick = "Pick"
Private Const tcFinish = "Finish"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const SDOCLINE = 1
Private Const SDOCNO = 2
Private Const SITMCODE = 3
Private Const SITMNAME = 4
Private Const SVDRCODE = 5
Private Const SQTY = 6
Private Const SOUTQTY = 7
Private Const SREM = 8
Private Const SAPRFLG = 9
Private Const SDUMMY = 10
Private Const SID = 11



Private Sub cboStaffNo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    
        
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
        
        tblDetail.SetFocus
    End If
End Sub


Private Sub cboStaffNo_LostFocus()
    FocusMe cboStaffNo, True
End Sub



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

    chk_cboDocNoFr = False
    
 If Chk_TrnHdDocNo("SO", cboDocNoFr, "") = False Then
        gsMsg = "Job No Not Exist!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNoFr.SetFocus
        Exit Function
  End If
  
  Get_RefDoc
    
  chk_cboDocNoFr = True
End Function

Private Function chk_cboStaffNo() As Boolean
Dim wsName As String

 chk_cboStaffNo = False
    
 If Chk_Salesman(cboStaffNo.Text, wlStaffID, wsName) = False Then
        gsMsg = "Worker Not Exist!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboStaffNo.SetFocus
        Exit Function
  End If
  
  lblDspName.Caption = wsName
    
  chk_cboStaffNo = True
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
           Call cmdSave
        
        If tbrProcess.Buttons(tcPick).Enabled = False Then Exit Sub
         '  Call cmdPick
        
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
             
        Case vbKeyF5
           Call cmdSelect(1)
           
        Case vbKeyF6
           Call cmdSelect(0)
        
        
            
    End Select
End Sub







Private Sub tabDetailInfo_Click(PreviousTab As Integer)



Call cmdRefresh


End Sub



Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    
   If tbrProcess.Buttons(Button.Key).Enabled = False Then Exit Sub
        
    
    Select Case Button.Key
        Case tcConvert
            Call cmdSave
                
        Case tcPick
         '  Call cmdPick
        
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



     Call LoadRecord
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    Set waScrItm = Nothing
 '   Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmMRP002 = Nothing
 
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
 '   wsFormID = "MRP002"
End Sub



Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
  '  Call Get_Scr_Item("TOOLTIP_A", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblStaffNo.Caption = Get_Caption(waScrItm, "StaffNo")
    lblJobRef.Caption = Get_Caption(waScrItm, "JOBREF")
    
                
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SVDRCODE).Caption = Get_Caption(waScrItm, "SVDRCODE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SOUTQTY).Caption = Get_Caption(waScrItm, "SOUTQTY")
        .Columns(SREM).Caption = Get_Caption(waScrItm, "SREM")
        .Columns(SAPRFLG).Caption = Get_Caption(waScrItm, "SAPRFLG")
        
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "OPT1")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "OPT2")
    
    
    
    With tbrProcess
    .Buttons(tcConvert).ToolTipText = Get_Caption(waScrItm, tcConvert) & "(F10)"
    .Buttons(tcPick).ToolTipText = Get_Caption(waScrItm, tcPick)
    .Buttons(tcFinish).ToolTipText = Get_Caption(waScrItm, tcFinish)
    
    
    .Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F7)"
    .Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F11)"
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
            Case SVDRCODE
            
              If Chk_grdVdrCode(.Columns(ColIndex).Text, .Columns(SITMCODE).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
              End If
                
            Case SREM
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
0
   Dim wsSQL As String
    Dim wiTop As Long
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err
    
    With tblDetail
        Select Case ColIndex
            Case SVDRCODE
                
            wsSQL = "SELECT VDRCODE, VDRNAME "
            wsSQL = wsSQL & " FROM mstVENDOR, mstVDRITEM, MSTITEM "
            wsSQL = wsSQL & " WHERE ITMCODE = '" & Set_Quote(.Columns(SITMCODE).Text) & "' "
            wsSQL = wsSQL & " AND VDRID = VDRITEMVDRID "
            wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
            wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
            wsSQL = wsSQL & " ORDER BY VDRCODE "
            
                
            Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLVDRCODE", Me.Width, Me.Height)
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
            Case SREM
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
                Case SAPRFLG
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
        
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                
                 
                Case SVDRCODE
                 
                 Call Chk_grdVdrCode(waResult(LastRow, SVDRCODE), waResult(LastRow, SITMCODE))
                 
                Case SREM
                 
                 Call Chk_grdQty(waResult(LastRow, SREM))
                 
            End Select
        End If
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
        .MultipleLines = 0
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
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 30
               Case SITMNAME
                   .Columns(wiCtr).Width = 2500
                   .Columns(wiCtr).DataWidth = 60
                Case SVDRCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = False
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
                Case SAPRFLG
                    .Columns(wiCtr).Width = 600
                    .Columns(wiCtr).DataWidth = 15
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
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
    Dim wiCtr As Integer
    Dim wiLastNo As Integer
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    Call Set_tbrProcess
    
    Select Case tabDetailInfo.Tab
    Case 0
    
    wsSQL = "SELECT SODTID DTID, SOPTJDOCLINE DOCLINE, SODTDOCLINE DOCNO, ITMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITMNAME , ITMITMTYPECODE, "
    wsSQL = wsSQL & " SODTTOTQTY QTY, SODTSCHQTY OUTQTY, SODTUPDFLG APRFLG, VDRCODE "
    wsSQL = wsSQL & " FROM  SOASOHD, SOASOPTJ, SOASODT, MSTITEM, mstVendor "
    wsSQL = wsSQL & " WHERE SOHDDOCNO = '" & cboDocNoFr & "' "
    wsSQL = wsSQL & " AND SOHDDOCID = SOPTJDOCID "
    wsSQL = wsSQL & " AND SOPTJID = SODTPTJID "
    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
    wsSQL = wsSQL & " AND SODTVDRID = VDRID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '1'"
    wsSQL = wsSQL & " AND (SODTTOTQTY-SODTSCHQTY) > 0 "
    wsSQL = wsSQL & " ORDER BY SOPTJDOCLINE, SODTDOCLINE "
    
    Case 1
    
    wsSQL = "SELECT PODTID DTID, PODTDOCLINE DOCLINE, POHDDOCNO DOCNO, ITMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITMNAME, ITMITMTYPECODE, "
    wsSQL = wsSQL & " PODTQTY QTY, PODTRECQTY OUTQTY, PODTUPDFLG APRFLG, VDRCODE "
    wsSQL = wsSQL & " FROM POPPOHD, POPPODT, SOASOHD, mstITEM, mstVendor "
    wsSQL = wsSQL & " WHERE SOHDDOCNO = '" & cboDocNoFr & "'"
    wsSQL = wsSQL & " AND POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND POHDREFDOCID = SOHDDOCID "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND POHDVDRID = VDRID "
  '  wsSql = wsSql & " AND (PODTQTY-PODTRECQTY) > 0 "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO, PODTDOCLINE "
     
    
    End Select
    
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
    
    wiCtr = 0
    wiLastNo = 0
    
     
    With waResult
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
     .AppendRows
        
        If wiLastNo <> To_Value(ReadRs(rsRcd, "DOCLINE")) Then
        wiCtr = 1
        Else
        wiCtr = wiCtr + 1
        End If
        wiLastNo = To_Value(ReadRs(rsRcd, "DOCLINE"))
        
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCLINE) = Format(ReadRs(rsRcd, "DOCLINE"), "000")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "ITMNAME")
        waResult(.UpperBound(1), SVDRCODE) = ReadRs(rsRcd, "VDRCODE")
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), SOUTQTY) = Format(To_Value(ReadRs(rsRcd, "OUTQTY")), gsAmtFmt)
        waResult(.UpperBound(1), SREM) = Format(To_Value(ReadRs(rsRcd, "QTY")) - To_Value(ReadRs(rsRcd, "OUTQTY")), gsQtyFmt)
        waResult(.UpperBound(1), SAPRFLG) = ReadRs(rsRcd, "APRFLG")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DTID")
        
    
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
        
        
        If Chk_grdVdrCode(waResult(LastRow, SVDRCODE), waResult(LastRow, SITMCODE)) = False Then
            .Col = SVDRCODE
            .Row = LastRow
            Exit Function
        End If
        
        
        If Chk_grdQty(waResult(LastRow, SREM)) = False Then
                .Col = SREM
               Exit Function
        End If
        
        If wiActFlg = 1 Then
        
        If waResult(LastRow, SAPRFLG) = "Y" Then
            .Col = SAPRFLG
             gsMsg = "已採購!不能轉移採購單"
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
    
    If wiActFlg = 2 Then
    If chk_cboStaffNo = False Then
        Exit Function
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

Private Sub cmdSave()

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wsTrnCd As String
    Dim wsDteTim As String
    Dim wiRet As Integer
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    wsDteTim = Change_SQLDate(Now)
    
    wiActFlg = 1
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    
    gsMsg = "你是否確認要轉換採購單?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

 
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
      If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_MRP002A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, wsDteTim)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SVDRCODE))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, SREM))
                Call SetSPPara(adcmdSave, 5, wsFormID)
                Call SetSPPara(adcmdSave, 6, gsUserID)
                Call SetSPPara(adcmdSave, 7, wsGenDte)
                adcmdSave.Execute
                wiRet = GetSPPara(adcmdSave, 8)
                
                If wiRet < 0 Then
                wsDocNo = waResult(wiCtr, SITMCODE)
                GoTo USP_MRP002A_Err
                End If
                
            End If
        Next
    End If
    
    
    
    adcmdSave.CommandText = "USP_MRP002B"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
     
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, wsDteTim)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, gsLangID)
    Call SetSPPara(adcmdSave, 5, wsFormID)
    Call SetSPPara(adcmdSave, 6, wsGenDte)
    adcmdSave.Execute
    wsDocNo = GetSPPara(adcmdSave, 7)
    
     
    cnCon.CommitTrans
    
  
    gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
        

    
    Set adcmdSave = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault


    Exit Sub
        
    
USP_MRP002A_Err:

    gsMsg = "物料 " & wsDocNo & " 供應商價格資料不足夠!不能採購"
    MsgBox gsMsg, vbOKOnly, gsTitle
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing

    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub


Private Sub Set_tbrProcess()

With tbrProcess
    
    Select Case tabDetailInfo.Tab
    Case 0
    
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcPick).Enabled = True
     Case 1
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcPick).Enabled = False
     
    End Select
    
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    

    
End With

End Sub

Private Function Chk_grdVdrCode(ByVal inCode As String, ByVal inItmCode As String) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset

    
    If Trim(inCode) = "" Then
        gsMsg = "沒有輸入供應商!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If


    wsSQL = "SELECT VdrItemCurr, VdrItemCost, VdrItemUomCode FROM MstVendor, mstVdrItem, MstItem "
    wsSQL = wsSQL & " WHERE VdrCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & " And ItmCode = '" & Set_Quote(inItmCode) & "' "
    wsSQL = wsSQL & " And VdrItemVdrID = VdrID "
    wsSQL = wsSQL & " And VdrItemItmID = ItmID "
    wsSQL = wsSQL & " And VdrItemStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_grdVdrCode = True
    Else
        gsMsg = "沒有此供應商價格!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdVdrCode = False
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
    
   
    
End Function
