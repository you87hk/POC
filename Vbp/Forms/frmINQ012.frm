VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmINQ012 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmINQ012.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   1920
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmINQ012.frx":0442
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   11280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":341F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":3CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":414B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":459D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":48B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":4D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":578F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":5BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":64BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":67E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":6C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":6F55
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":7271
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":76C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":79E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":7CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":8151
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ012.frx":846D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   11775
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label lblJobRef 
         Caption         =   "CUSNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label lblDspJobRef1 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   585
         Width           =   5415
      End
      Begin VB.Label lblDspJobRef2 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   945
         Width           =   5415
      End
      Begin VB.Label lblDspJobRef3 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   1305
         Width           =   5415
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "frmINQ012.frx":8791
      TabIndex        =   0
      Top             =   2280
      Width           =   11775
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Print"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   3
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmINQ012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Private wbErr As Boolean


Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer
Private wsDteTim As String

Private wlKey As Long

Private Const tcPrint = "Print"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"


Private Const SDOCDATE = 0
Private Const STRNCODE = 1
Private Const SDOCNO = 2
Private Const SCUSCODE = 3
Private Const SCUSNAME = 4
Private Const SUPDUSR = 5
Private Const SUPDDATE = 6
Private Const sStatus = 7
Private Const SID = 8
Private Const SDUMMY = 9



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
        tblDetail.SetFocus
        
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



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
       
            
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
            
     
        Case vbKeyF7
            Call LoadRecord
        
      
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


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, SDOCDATE, SID
  
    
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
            Case "ComboBox"
                MyControl.Clear
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
    
    
    wlKey = 0
    cboDocNoFr.Text = ""
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   
   
    cnCon.Execute "DELETE FROM RPTINQ012 WHERE RPTUSRID = '" & gsUserID & "' AND RPTDTETIM = '" & wsDteTim & "' "
    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmINQ012 = Nothing


    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "INQ012"

    
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "JOBNO")
    lblJobRef.Caption = Get_Caption(waScrItm, "JOBREF")
       
    
    
    
    With tblDetail
        .Columns(STRNCODE).Caption = Get_Caption(waScrItm, "STRNCODE")
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCUSCODE).Caption = Get_Caption(waScrItm, "SCUSCODE")
        .Columns(SCUSNAME).Caption = Get_Caption(waScrItm, "SCUSNAME")
        .Columns(SUPDUSR).Caption = Get_Caption(waScrItm, "SUPDUSR")
        .Columns(SUPDDATE).Caption = Get_Caption(waScrItm, "SUPDDATE")
        .Columns(sStatus).Caption = Get_Caption(waScrItm, "SSTATUS")
        
        
        
    End With
    
    
    
    
'    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F11)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    

End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .Update
    End With
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    
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





Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode

            
        Case vbKeyReturn
            Select Case .Col
            Case sStatus
                 KeyCode = vbKeyDown
                 .Col = SDOCDATE
            Case Else
                 KeyCode = vbDefault
                 .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SDOCDATE Then
                .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case sStatus
                KeyCode = vbKeyDown
                    .Col = SDOCDATE
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
                 
                Case SCUSNAME
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SCUSNAME).Text
                    
                  
            End Select
        End If
    End With
    
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
            
        Case tcPrint
        
            Call cmdPrint
            
        Case tcCancel
        
           Call cmdCancel

            
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call LoadRecord
            
    End Select
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
        
        For wiCtr = SDOCDATE To SDUMMY
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case STRNCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 3
                Case SDOCDATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SDOCNO
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 20
                Case SCUSCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SCUSNAME
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 60
                Case SUPDUSR
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 20
                Case SUPDDATE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 10
                Case sStatus
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 10
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
    Dim wiCtr As Long


    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    wsDteTim = Change_SQLDate(Now)
    
    Call cmdSave
    
    wsSQL = "SELECT RPTDOCID, RPTDOCNO, RPTDOCDATE, RPTTRNCODE, RPTDOCNO, RPTCUSCODE, RPTCUSNAME, RPTUPDUSR, RPTUPDDATE, RPTSTATUS "
    wsSQL = wsSQL & " From RPTINQ012 "
    wsSQL = wsSQL & " WHERE RPTUSRID = '" & gsUserID & "' "
    wsSQL = wsSQL & " AND RPTDTETIM = '" & wsDteTim & "' "
    wsSQL = wsSQL & " ORDER BY RPTDOCDATE, RPTDOCNO "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SDOCDATE, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
    
     
    With waResult
    .ReDim 0, -1, SDOCDATE, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

       .AppendRows
        waResult(.UpperBound(1), STRNCODE) = ReadRs(rsRcd, "RPTTRNCODE")
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "RPTDOCDATE")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "RPTDOCNO")
        waResult(.UpperBound(1), SCUSCODE) = ReadRs(rsRcd, "RPTCUSCODE")
        waResult(.UpperBound(1), SCUSNAME) = ReadRs(rsRcd, "RPTCUSNAME")
        waResult(.UpperBound(1), SUPDUSR) = ReadRs(rsRcd, "RPTUPDUSR")
        waResult(.UpperBound(1), SUPDDATE) = ReadRs(rsRcd, "RPTUPDDATE")
        waResult(.UpperBound(1), sStatus) = ReadRs(rsRcd, "RPTSTATUS")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "RPTDOCID")
        
        'End If

      
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

Private Sub cmdPrint()
    
    
End Sub

Private Sub cmdSave()
    Dim adcmdSave As New ADODB.Command

     
    On Error GoTo cmdSave_Err
    
    'MousePointer = vbHourglass
    
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    adcmdSave.CommandText = "USP_RPTINQ012"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
     
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, wsDteTim)
    Call SetSPPara(adcmdSave, 3, wlKey)
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
