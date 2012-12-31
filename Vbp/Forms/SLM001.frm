VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmSLM001 
   BackColor       =   &H8000000A&
   Caption         =   "營業員資料"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "SLM001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "SLM001.frx":08CA
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboSaleCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "營業員"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   8355
      Begin VB.Frame FraSaleType 
         Caption         =   "MLTYPE"
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   5535
         Begin VB.OptionButton optSaleType 
            Caption         =   "PURCHASE"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optSaleType 
            Caption         =   "SALES"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox txtSaleCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Tag             =   "K"
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtSaleName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label lblDspSaleLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6360
         TabIndex        =   9
         Top             =   2400
         Width           =   1785
      End
      Begin VB.Label lblDspSaleLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2400
         TabIndex        =   8
         Top             =   2400
         Width           =   1785
      End
      Begin VB.Label lblSaleLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4320
         TabIndex        =   7
         Top             =   2445
         Width           =   1860
      End
      Begin VB.Label lblSaleLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   2445
         Width           =   1980
      End
      Begin VB.Label lblSaleCode 
         Caption         =   "編碼 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblSaleName 
         Caption         =   "姓名 :"
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   1020
         Width           =   1380
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SLM001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "開新視窗 (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "新增 (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "修改 (F5)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "刪除 (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "儲存 (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "尋找 (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmSLM001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wsFormCaption As String
Private waScrItm  As New XArrayDB
Private waScrToolTip As New XArrayDB
 
Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"

Private wiAction As Integer
Private wsActNam(4) As String

Private wlKey As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private Const wsKeyType = "MstSalesman"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboSaleCode_LostFocus()
    FocusMe cboSaleCode, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        Case vbKeyF6
            Call cmdOpen
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
            
        
        Case vbKeyF5
            If wiAction = DefaultPage Then Call cmdEdit
       
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF9
        
            If tbrProcess.Buttons(tcFind).Enabled = True Then
                Call cmdFind
            End If
            
        Case vbKeyF10
        
            If tbrProcess.Buttons(tcSave).Enabled = True Then
                Call cmdSave
            End If
            
        Case vbKeyF11
        
            If wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec Then Call cmdCancel
        
        Case vbKeyF12
        
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    Dim iCounter As Integer
    Dim iTabs As Integer
    Dim vToolTip As Variant
    
    MousePointer = vbHourglass
  
    wsFormCaption = Me.Caption
    
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 3855
        Me.Width = 8700
    End If
End Sub

'-- Set toolbar buttons status in different mode, Default, AddEdit, None.
Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = True
                .Buttons(tcEdit).Enabled = True
                .Buttons(tcDelete).Enabled = True
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
            
        Case "AfrActAdd"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "AfrActEdit"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = True
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "AfrKey"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
        
        
        Case "ReadOnly"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcFind).Enabled = True
                .Buttons(tcExit).Enabled = True
            
            End With
    End Select
End Sub

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            Me.txtSaleName.Enabled = False
            
            Me.cboSaleCode.Enabled = False
            Me.cboSaleCode.Visible = False
            Me.txtSaleCode.Visible = True
            Me.txtSaleCode.Enabled = False
            
            optSaleType(0).Enabled = False
            optSaleType(1).Enabled = False
            
            
        Case "AfrActAdd"
            Me.cboSaleCode.Enabled = False
            Me.cboSaleCode.Visible = False
            
            Me.txtSaleCode.Enabled = True
            Me.txtSaleCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboSaleCode.Enabled = True
            Me.cboSaleCode.Visible = True
            
            Me.txtSaleCode.Enabled = False
            Me.txtSaleCode.Visible = False
            
        Case "AfrKey"
            Me.txtSaleName.Enabled = True
            optSaleType(0).Enabled = True
            optSaleType(1).Enabled = True
            
            
            Me.cboSaleCode.Enabled = False
            Me.txtSaleCode.Enabled = False
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    If Chk_txtSaleName = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT MstSalesman.SaleCode, MstSalesman.* "
    wsSQL = wsSQL + "From MstSalesman "
    wsSQL = wsSQL + "WHERE (((MstSalesman.SaleCode)='" + Set_Quote(cboSaleCode) + "') "
    wsSQL = wsSQL + "AND ((MstSalesman.SaleStatus)='1'));"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "SaleID")
        
        Me.cboSaleCode = ReadRs(rsRcd, "SaleCode")
        Me.txtSaleName = ReadRs(rsRcd, "SaleName")
        Me.lblDspSaleLastUpd = ReadRs(rsRcd, "SaleLastUpd")
        Me.lblDspSaleLastUpdDate = ReadRs(rsRcd, "SaleLastUpdDate")
        SetSaleType ReadRs(rsRcd, "SaleType")
        
        LoadRecord = True
    End If
    
    rsRcd.Close
    
    Set rsRcd = Nothing
    
End Function

Private Sub Form_Unload(Cancel As Integer)
   If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    

    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set frmSLM001 = Nothing
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case tcOpen
            Call cmdOpen
        
        Case tcAdd
            Call cmdNew
            
        Case tcEdit
            Call cmdEdit
        
        Case tcDelete
            
            Call cmdDel
            
        Case tcSave
            
            Call cmdSave
            
        Case tcCancel
            
            If tbrProcess.Buttons(tcSave).Enabled = True Then
                gsMsg = "你是否確定儲存現時之變更而離開?"
                If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
                    Call cmdCancel
                End If
            Else
                Call cmdCancel
            End If
        
        Case tcFind
            
            Call OpenPromptForm
            
        Case tcExit
        
            Unload Me
            
    End Select
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
 '   Me.Left = 0
 '   Me.Top = 0
 '   Me.Width = Screen.Width
 '   Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "SLM001"
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
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

    
    wiAction = DefaultPage
    wlKey = 0
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    tblCommon.Visible = False
    Me.Caption = wsFormCaption
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtSaleCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboSaleCode.SetFocus
       
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboSaleCode.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, cboSaleCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtSaleName.SetFocus
End Sub

Private Function Chk_SaleCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_SaleCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT SaleStatus "
    wsSQL = wsSQL & " FROM MstSalesman WHERE SaleCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "SaleStatus")
    Chk_SaleCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboSaleCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboSaleCode = False
    
    If Trim(cboSaleCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        Exit Function
    End If
    
    If Chk_SaleCode(cboSaleCode.Text, wsStatus) = False Then
        gsMsg = "營業員編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "營業員編碼已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboSaleCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboSaleCode = True
End Function

Private Function Chk_txtSaleCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtSaleCode = False
    
    If Trim(txtSaleCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtSaleCode.SetFocus
        Exit Function
    End If
    
    If Chk_SaleCode(txtSaleCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "營業員編碼已存在但已無效!"
        Else
            gsMsg = "營業員編碼已存在!"
        End If
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtSaleCode.SetFocus
        Exit Function
    End If
    Chk_txtSaleCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmSLM001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show
End Sub

Private Sub cmdNew()
    wiAction = AddRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdEdit()
    wiAction = CorRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdFind()
     Call OpenPromptForm
End Sub

Private Sub cmdDel()
    wiAction = DelRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdCancel()
    If tbrProcess.Buttons(tcSave).Enabled = True Then
        Select Case wiAction
            Case AddRec
                Call Ini_Scr
                Call cmdNew
                
            Case CorRec
                Call UnLockAll(wsConnTime, wsFormID)
                Call Ini_Scr
                Call cmdEdit
                
            Case DelRec
                Call UnLockAll(wsConnTime, wsFormID)
                Call Ini_Scr
                Call cmdDel
        End Select
    Else
        Call Ini_Scr
    End If
End Sub

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboSaleCode, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If wiAction = DelRec Then
        If MsgBox("你是否確定要刪除此記錄?", vbYesNo, gsTitle) = vbNo Then
            cmdCancel
            MousePointer = vbDefault
            Exit Function
        End If
    Else
        If InputValidation() = False Then
            MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SLM001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, txtSaleCode, cboSaleCode))
    Call SetSPPara(adcmdSave, 4, txtSaleName)
    Call SetSPPara(adcmdSave, 5, GetSaleType())
    Call SetSPPara(adcmdSave, 6, gsUserID)
    Call SetSPPara(adcmdSave, 7, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 8)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - SLM001!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
        gsMsg = "已成功儲存!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    End If
    
    Call cmdCancel
    
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

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
    SaveData = False
    
    If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) Then
        gsMsg = "你是否確定要儲存現時之作業?"
        If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
            Exit Function
        Else
                If cmdSave = True Then
                    Exit Function
                End If
        End If
        SaveData = True
    Else
        SaveData = False
    End If
    
End Function

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim sSQL As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "營業員編碼"
    vFilterAry(1, 2) = "SaleCode"
    
    vFilterAry(2, 1) = "註解"
    vFilterAry(2, 2) = "StoreDesc"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "營業員編碼"
    vAry(1, 2) = "StoreCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    vAry(2, 2) = "StoreDesc"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstSalesman.SaleCode, MstSalesman.SaleName "
        sSQL = sSQL + "FROM MstSalesman "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstSalesman.SaleStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstSalesman.SaleCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboSaleCode Then
        cboSaleCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtSaleCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtSaleCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtSaleCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtSaleCode_LostFocus()
    FocusMe txtSaleCode, True
End Sub

Private Sub txtSaleName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtSaleName, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtSaleName = True Then
        
            Call Opt_Setfocus(optSaleType, 2, 0)
            
        End If
        
        
    End If
End Sub

Private Sub txtSaleCode_GotFocus()
    FocusMe txtSaleCode
End Sub

Private Sub txtSaleName_GotFocus()
    FocusMe txtSaleName
End Sub

Private Function Chk_txtSaleName() As Boolean
    Chk_txtSaleName = False
    
    If Trim(txtSaleName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtSaleName.SetFocus
        Exit Function
    End If
    
    Chk_txtSaleName = True
End Function

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

Private Sub cboSaleCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboSaleCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboSaleCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode
    
    wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleStatus = '1'"
    wsSQL = wsSQL & " AND SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
   
    wsSQL = wsSQL & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboSaleCode.Left, cboSaleCode.Top + cboSaleCode.Height, tblCommon, "SLM001", "TBLSLMCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboSaleCode_GotFocus()
    FocusMe cboSaleCode
End Sub

Private Sub txtSaleName_LostFocus()
    FocusMe txtSaleName, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT SaleStatus FROM MstSalesman WHERE SaleCode = '" & Set_Quote(txtSaleCode) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_KeyExist = True
    Else
        Chk_KeyExist = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "SaleCode"
        .KeyLen = 10
        Set .ctlKey = txtSaleCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblSaleName.Caption = Get_Caption(waScrItm, "SALENAME")
    lblSaleLastUpd.Caption = Get_Caption(waScrItm, "SALELASTUPD")
    lblSaleLastUpdDate.Caption = Get_Caption(waScrItm, "SALELASTUPDDATE")
    
    optSaleType(0).Caption = Get_Caption(waScrItm, "OPT0")
    optSaleType(1).Caption = Get_Caption(waScrItm, "OPT1")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "SLMADD")
    wsActNam(2) = Get_Caption(waScrItm, "SLMEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SLMDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub SetSaleType(ByVal inCode As String)
    Select Case inCode
        Case "S"
            optSaleType(0).Value = True
            
        Case "W"
            optSaleType(1).Value = True
            
        
    End Select
End Sub

Private Function GetSaleType() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 1
        If optSaleType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetSaleType = "S"
            
        Case 1
            GetSaleType = "W"
        
        
    End Select
End Function
