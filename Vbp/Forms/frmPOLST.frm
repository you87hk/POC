VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPOLST 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "快速搜尋"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmPOLST.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   7406.108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   10061.66
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   8880
      OleObjectBlob   =   "frmPOLST.frx":0442
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboVdrNoFr 
      Height          =   300
      Left            =   1920
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.ComboBox cboVdrNoTo 
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5400
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   9735
      Begin MSMask.MaskEdBox medPrdTo 
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPrdFr 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblVdrNoFr 
         Caption         =   "Vendor Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblPrdFr 
         Caption         =   "Period From"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   990
         Width           =   1890
      End
      Begin VB.Label lblVdrNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   13
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblPrdTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   12
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lblDocNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   11
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   1890
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   9600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":341F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":3CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":414B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":459D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":48B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":4D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":578F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":5BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":64BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":67E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":6C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPOLST.frx":6F5D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmPOLST.frx":7281
      TabIndex        =   6
      Top             =   1920
      Width           =   9855
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   16
      Top             =   7080
      Width           =   9855
   End
End
Attribute VB_Name = "frmPOLST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Dim waInvDoc As New XArrayDB
Private wbErr As Boolean

Private wsCurr As String
Private wdExcr As Double
Private wlVdrID As Long
Private wlDocID As Long
Private wsTrnCd As String
Private wlLineNo As Long

Private wiExit As Boolean
Private wiUpdate As Boolean

Private wsFormCaption As String
Private wsFormID As String


Private Const tcGo = "Go"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"

Private Const SSEL = 0
Private Const SDOCNO = 1
Private Const sItmCode = 2
Private Const SITMDESC = 3
Private Const SWHSCODE = 4
Private Const SLOTNO = 5
Private Const SUPRICE = 6
Private Const SQTY = 7
Private Const SAMT = 8
Private Const SNET = 9
Private Const SDUMMY = 10
Private Const SID = 11

Private Const LINENO = 0
Private Const PONO = 1
Private Const BOOKCODE = 2
Private Const BARCODE = 3
Private Const WhsCode = 4
Private Const LOTNO = 5
Private Const BOOKNAME = 6
Private Const PUBLISHER = 7
Private Const Qty = 8
Private Const Price = 9
Private Const DisPer = 10
Private Const Dis = 11
Private Const Amt = 12
Private Const Net = 13
Private Const Netl = 14
Private Const Disl = 15
Private Const Amtl = 16
Private Const BOOKID = 17
Private Const POID = 18

Public Property Get InvDoc() As XArrayDB
    Set InvDoc = waInvDoc
End Property

Public Property Let InvDoc(inInvDoc As XArrayDB)
    Set waInvDoc = inInvDoc
End Property

Public Property Get InLineNo() As Long
     InLineNo = wlLineNo
End Property

Public Property Let InCurr(inDocCurr As String)
     wsCurr = inDocCurr
End Property

Public Property Let inExcr(inDocExcr As Double)
     wdExcr = inDocExcr
End Property

Public Property Let InVdrID(inDocVdr As Long)
     wlVdrID = inDocVdr
End Property

Public Property Let InLineNo(inLine As Long)
     wlLineNo = inLine
End Property

Public Property Let InDocID(inDoc As Long)
     wlDocID = inDoc
End Property

Public Property Let InTrnCd(inTrn As String)
     wsTrnCd = inTrn
End Property

Private Sub cboVdrNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND VdrStatus <> '2' "
    wsSql = wsSql & " ORDER BY Vdrcode "
    Call Ini_Combo(2, wsSql, cboVdrNoFr.Left, cboVdrNoFr.Top + cboVdrNoFr.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrNoFr_GotFocus()
        FocusMe cboVdrNoFr
    Set wcCombo = cboVdrNoFr
End Sub

Private Sub cboVdrNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboVdrNoFr.Text) <> "" And _
            Trim(cboVdrNoTo.Text) = "" Then
            cboVdrNoTo.Text = cboVdrNoFr.Text
        End If
        cboVdrNoTo.SetFocus
    End If
End Sub

Private Sub cboVdrNoFr_LostFocus()
    FocusMe cboVdrNoFr, True
End Sub

Private Sub cboVdrNoTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND VdrStatus <> '2' "
    wsSql = wsSql & " ORDER BY Vdrcode "
    Call Ini_Combo(2, wsSql, cboVdrNoTo.Left, cboVdrNoTo.Top + cboVdrNoTo.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrNoTo_GotFocus()
    FocusMe cboVdrNoTo
    Set wcCombo = cboVdrNoTo
End Sub

Private Sub cboVdrNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrNoTo = False Then
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub

Private Sub cboVdrNoTo_LostFocus()
    FocusMe cboVdrNoTo, True
End Sub

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        
        If LoadRecord = True Then
            tblDetail.SetFocus
        End If
       
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub

Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub

Private Sub cboDocNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSql = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSql = wsSql & " FROM popPOHD, popPODT, MstVendor "
    wsSql = wsSql & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSql = wsSql & " AND POHDVDRID  = VDRID "
    wsSql = wsSql & " AND POHDDOCID  = PODTDOCID "
    wsSql = wsSql & " AND POHDSTATUS = '1' "
    wsSql = wsSql & " GROUP BY POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSql = wsSql & " HAVING SUM(PODTBALQTY) <> 0 "
    wsSql = wsSql & " ORDER BY POHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
        If Trim(cboDocNoFr.Text) <> "" And _
            Trim(cboDocNoTo.Text) = "" Then
            cboDocNoTo.Text = cboDocNoFr.Text
        End If
        cboDocNoTo.SetFocus
    End If
End Sub

Private Sub cboDocNoFr_LostFocus()
    FocusMe cboDocNoFr, True
End Sub

Private Sub cboDocNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    wsSql = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSql = wsSql & " FROM popPOHD, popPODT, MstVendor "
    wsSql = wsSql & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSql = wsSql & " AND POHDVDRID  = VDRID "
    wsSql = wsSql & " AND POHDDOCID  = PODTDOCID "
    wsSql = wsSql & " AND POHDSTATUS = '1' "
    wsSql = wsSql & " GROUP BY POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSql = wsSql & " HAVING SUM(PODTBALQTY) <> 0 "
    wsSql = wsSql & " ORDER BY POHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Call cboDocNoTo_GotFocus
            Exit Sub
        End If
        
       cboVdrNoFr.SetFocus

    End If
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
End Sub

Private Function chk_cboDocNoTo() As Boolean
    chk_cboDocNoTo = False
    
    If UCase(cboDocNoFr.Text) > UCase(cboDocNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboDocNoTo = True
End Function

Private Function chk_cboVdrNoTo() As Boolean
    chk_cboVdrNoTo = False
    
    If UCase(cboVdrNoFr.Text) > UCase(cboVdrNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboVdrNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboVdrNoTo = True
End Function

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        gsMsg = "To must > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
           Call cmdOK
            
        Case vbKeyF3
           Call cmdCancel
            
        Case vbKeyF12
            Me.Hide
        
        Case vbKeyF5
            Call LoadRecord
        
        Case vbKeyF9
           Call cmdSelect(1)
           
        Case vbKeyF10
           Call cmdSelect(0)
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

Private Sub cmdOK()
    MousePointer = vbHourglass
    Unload Me
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
    wiUpdate = True
    
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    
    cboDocNoFr.Text = ""
    cboDocNoTo.Text = ""
    cboVdrNoFr.Text = Get_TableInfo("MSTVENDOR", "VDRID = " & wlVdrID, "VDRCODE")
    cboVdrNoTo.Text = cboVdrNoFr.Text
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If wiExit = False Then
     
       Cancel = True
       If wiUpdate = True Then Call UpdateRecord
       wiExit = True
       Me.Hide
       Exit Sub
    End If
    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set waInvDoc = Nothing
    
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "POLST"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblVdrNoFr.Caption = Get_Caption(waScrItm, "VDRNOFR")
    lblVdrNoTo.Caption = Get_Caption(waScrItm, "VDRNOTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(sItmCode).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMDESC).Caption = Get_Caption(waScrItm, "SITMDESC")
        .Columns(SWHSCODE).Caption = Get_Caption(waScrItm, "SWHSCODE")
        .Columns(SLOTNO).Caption = Get_Caption(waScrItm, "SLOTNO")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SAMT).Caption = Get_Caption(waScrItm, "SAMT")
        .Columns(SNET).Caption = Get_Caption(waScrItm, "SNET")
    End With
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F2)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F9)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F10)"

End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .Update
    End With
End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
       
        Case vbKeyReturn
            Select Case .Col
                Case SNET
                    KeyCode = vbKeyDown
                    .Col = SSEL
                Case SSEL, SDOCNO, SWHSCODE, SLOTNO, SUPRICE, SQTY, SAMT
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case sItmCode
                    KeyCode = vbDefault
                    .Col = SWHSCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col = SWHSCODE Then
                   .Col = sItmCode
            ElseIf .Col <> SSEL Then
                   .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SNET
                    KeyCode = vbKeyDown
                    .Col = SSEL
                Case SSEL, SDOCNO, SWHSCODE, SLOTNO, SUPRICE, SQTY, SAMT
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case sItmCode
                    KeyCode = vbDefault
                    .Col = SWHSCODE
            End Select
        
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
       
    lblDspItmDesc.Caption = ""
    lblDspItmDesc.Caption = tblDetail.Columns(SITMDESC).Text
        
    Exit Sub

RowColChange_Err:
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Go"
            Call cmdOK
            
        Case "Cancel"
        
           Call cmdCancel
            
        Case "Exit"
           wiUpdate = False
           Unload Me
            
        Case "Refresh"
            Call LoadRecord
            
        Case "SAll"
        
           Call cmdSelect(1)
        
        Case "DAll"
        
           Call cmdSelect(0)
            
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then wcCombo.SetFocus

End Sub

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SSEL To SID
            .Columns(wiCtr).AllowSizing = False
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
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 1500
                Case sItmCode
                   .Columns(wiCtr).Width = 1500
                   .Columns(wiCtr).DataWidth = 13
                Case SITMDESC
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Visible = False
                Case SWHSCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SLOTNO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 20
                Case SQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case SUPRICE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SAMT
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SNET
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
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
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wdQty As Double
    Dim wdPrice As Double
    Dim wdAmt As Double
    Dim wdDis As Double
    Dim wdNet As Double
    Dim wdBalQty As Double
    
    LoadRecord = False
    Me.MousePointer = vbHourglass
    
    
    wsSql = "SELECT  POHDDOCNO, ITMCODE, PODTITEMID, PODTWHSCODE, PODTLOTNO, PODTITEMDESC ITNAME, PODTQTY, PODTBALQTY, PODTUPRICE, "
    wsSql = wsSql & " PODTDISPER, PODTAMT, PODTNET, PODTID, PODTDOCID "
    wsSql = wsSql & "FROM popPOHD, popPODT, MstVendor, mstITEM "
    wsSql = wsSql & "WHERE POHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSql = wsSql & "AND VDRCODE BETWEEN '" & cboVdrNoFr & "' AND '" & IIf(Trim(cboVdrNoTo.Text) = "", String(10, "z"), Set_Quote(cboVdrNoTo.Text)) & "'"
    wsSql = wsSql & "AND POHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSql = wsSql & "AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSql = wsSql & "AND POHDDOCID = PODTDOCID "
    wsSql = wsSql & "AND POHDVDRID = VDRID "
    wsSql = wsSql & "AND PODTITEMID = ITMID "
    wsSql = wsSql & "AND POHDSTATUS = '1'"
'        wsSql = wsSql & "AND SODTBALQTY > 0 "
    wsSql = wsSql & "ORDER BY PODTDOCID, PODTDOCLINE "
    
     rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Me.MousePointer = vbNormal
        Set rsRcd = Nothing
        Exit Function
    End If
    
    With waResult
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
        
        wdBalQty = Get_PoBalQty(wsTrnCd, wlDocID, ReadRs(rsRcd, "PODTDOCID"), ReadRs(rsRcd, "PODTITEMID"), ReadRs(rsRcd, "PODTWHSCODE"), ReadRs(rsRcd, "PODTLOTNO"))

        .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "POHDDOCNO")
        waResult(.UpperBound(1), sItmCode) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SITMDESC) = ReadRs(rsRcd, "ITNAME")
        waResult(.UpperBound(1), SWHSCODE) = ReadRs(rsRcd, "PODTWHSCODE")
        waResult(.UpperBound(1), SLOTNO) = ReadRs(rsRcd, "PODTLOTNO")
        waResult(.UpperBound(1), SUPRICE) = Format(To_Value(ReadRs(rsRcd, "PODTUPRICE")), gsAmtFmt)
        waResult(.UpperBound(1), SQTY) = Format(wdBalQty, gsQtyFmt)
        If wdBalQty = To_Value(ReadRs(rsRcd, "PODTQTY")) Then
            waResult(.UpperBound(1), SAMT) = Format(To_Value(ReadRs(rsRcd, "PODTAMT")), gsAmtFmt)
            waResult(.UpperBound(1), SNET) = Format(To_Value(ReadRs(rsRcd, "PODTNET")), gsAmtFmt)
        Else
            wdPrice = To_Value(ReadRs(rsRcd, "PODTUPRICE"))
            wdQty = wdBalQty
            wdAmt = Format(wdPrice * wdQty, gsAmtFmt)
            wdDis = Format(wdAmt * To_Value(ReadRs(rsRcd, "PODTDISPER")) / 100, gsAmtFmt)
            wdNet = wdAmt - wdDis
            waResult(.UpperBound(1), SAMT) = Format(wdAmt, gsAmtFmt)
            waResult(.UpperBound(1), SNET) = Format(wdNet, gsAmtFmt)
        End If
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "PODTID")
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

Private Function UpdateRecord() As Boolean
    Dim wiCtr As Long
    Dim wsBarCode As String
    Dim wdPrice As Double
    Dim wsPub As String
    Dim wsBookID As String
    Dim wdDisPer As Double
    Dim wdQty As Double
    Dim wdDis As Double
    Dim wdAmt As Double
    Dim wdNet As Double
    Dim wsLotNo As String
    Dim wsPoId As String
    
    UpdateRecord = False
    
    With waInvDoc
          
    If waResult.UpperBound(1) >= 0 Then
         
    For wiCtr = 0 To waResult.UpperBound(1)
         If Trim(waResult(wiCtr, SSEL)) = "-1" Then
            If Trim(waResult(wiCtr, SID)) = "" Then Exit For
             If get_POInfo(waResult(wiCtr, SID), wsBookID, wsBarCode, wsPub, wdPrice, wdQty, wdDisPer, wdAmt, wdDis, wdNet, wsPoId) = False Then Exit For
             .AppendRows
             waInvDoc(.UpperBound(1), LINENO) = wlLineNo
             waInvDoc(.UpperBound(1), PONO) = waResult(wiCtr, SDOCNO)
             waInvDoc(.UpperBound(1), BOOKCODE) = waResult(wiCtr, sItmCode)
             waInvDoc(.UpperBound(1), BARCODE) = wsBarCode
             waInvDoc(.UpperBound(1), BOOKNAME) = waResult(wiCtr, SITMDESC)
             waInvDoc(.UpperBound(1), PUBLISHER) = wsPub
             waInvDoc(.UpperBound(1), WhsCode) = waResult(wiCtr, SWHSCODE)
             waInvDoc(.UpperBound(1), LOTNO) = waResult(wiCtr, SLOTNO)
             waInvDoc(.UpperBound(1), Price) = Format(wdPrice, gsAmtFmt)
             waInvDoc(.UpperBound(1), Qty) = Format(wdQty, gsQtyFmt)
             waInvDoc(.UpperBound(1), DisPer) = Format(wdDisPer, gsAmtFmt)
             waInvDoc(.UpperBound(1), Dis) = Format(wdDis, gsAmtFmt)
             waInvDoc(.UpperBound(1), Disl) = Format(wdDis * wdExcr, gsAmtFmt)
             waInvDoc(.UpperBound(1), Amt) = Format(wdAmt, gsAmtFmt)
             waInvDoc(.UpperBound(1), Amtl) = Format(wdAmt * wdExcr, gsAmtFmt)
             waInvDoc(.UpperBound(1), Net) = Format(wdNet, gsAmtFmt)
             waInvDoc(.UpperBound(1), Netl) = Format(wdNet * wdExcr, gsAmtFmt)
             waInvDoc(.UpperBound(1), BOOKID) = wsBookID
             waInvDoc(.UpperBound(1), POID) = wsPoId
             wlLineNo = wlLineNo + 1
         End If
    Next wiCtr
    End If
   
    End With
    
    UpdateRecord = True
End Function

Private Function get_POInfo(inID As String, outBookID As String, OutBarCode As String, outPub As String, outPrice As Double, outQty As Double, outDisPer As Double, outAmt As Double, outDis As Double, outNet As Double, outPoID As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim lstCurr As String
    Dim lstExcr As Double
    Dim lstPrice As Double
    Dim lstQty As Double
    Dim lstBalQty As Double
     
    get_POInfo = False

    wsSql = "SELECT POHDDOCID, POHDCURR, POHDEXCR, ITMID, ITMBARCODE, ITMPUBLISHER, "
    wsSql = wsSql & "PODTUPRICE, PODTQTY, PODTBALQTY, PODTDISPER, PODTAMT, "
    wsSql = wsSql & "PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL, PODTWHSCODE, PODTLOTNO "
    wsSql = wsSql & "FROM  POPPOHD, POPPODT, MstItem "
    wsSql = wsSql & "WHERE PODTID =  " & inID & " "
    wsSql = wsSql & "AND PODTITEMID = ITMID "
    wsSql = wsSql & "AND PODTDOCID = POHDDOCID "
  
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outPoID = ReadRs(rsRcd, "POHDDOCID")
    outBookID = ReadRs(rsRcd, "ITMID")
    OutBarCode = ReadRs(rsRcd, "ITMBARCODE")
    outPub = ReadRs(rsRcd, "ITMPUBLISHER")
    lstPrice = To_Value(ReadRs(rsRcd, "PODTUPRICE"))
    lstQty = To_Value(ReadRs(rsRcd, "PODTQTY"))
    lstBalQty = Get_PoBalQty(wsTrnCd, wlDocID, outPoID, ReadRs(rsRcd, "ITMID"), ReadRs(rsRcd, "PODTWHSCODE"), ReadRs(rsRcd, "PODTLOTNO"))
    lstCurr = ReadRs(rsRcd, "POHDCURR")
    lstExcr = To_Value(ReadRs(rsRcd, "POHDEXCR"))
    outDisPer = To_Value(ReadRs(rsRcd, "PODTDISPER"))
    outQty = lstBalQty
    If Trim(UCase(wsCurr)) = Trim(UCase(lstCurr)) And lstQty = lstBalQty Then
        outPrice = lstPrice
        outAmt = To_Value(ReadRs(rsRcd, "PODTAMT"))
        outDis = To_Value(ReadRs(rsRcd, "PODTDIS"))
        outNet = To_Value(ReadRs(rsRcd, "PODTNET"))
    Else
        outPrice = lstPrice * lstExcr / wdExcr
        outAmt = Format(outPrice * outQty, gsAmtFmt)
        outDis = Format(outAmt * outDisPer / 100, gsAmtFmt)
        outNet = Format(outAmt - outDis, gsAmtFmt)
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
  
    get_POInfo = True
    
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

