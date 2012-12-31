VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmAPW0011 
   BorderStyle     =   1  '單線固定
   Caption         =   "New Key"
   ClientHeight    =   5100
   ClientLeft      =   990
   ClientTop       =   2415
   ClientWidth     =   8970
   Icon            =   "frmAPW0011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8970
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboDocLine 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2610
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   7920
      OleObjectBlob   =   "frmAPW0011.frx":030A
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItmType 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   2370
   End
   Begin VB.ComboBox cboVdrCode 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   2370
   End
   Begin VB.ComboBox cboItmCode 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   2370
   End
   Begin VB.Frame fraHeader 
      Height          =   4080
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   8715
      Begin VB.TextBox txtQty 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtMU 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   735
         Left            =   7080
         Picture         =   "frmAPW0011.frx":2A0D
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   735
         Left            =   5520
         Picture         =   "frmAPW0011.frx":2D17
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtItmName 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   7005
      End
      Begin VB.Label lblQty 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label lblNet 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   1320
      End
      Begin VB.Label lblAmt 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Label lblMU 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1320
      End
      Begin VB.Label lblUnitPrice 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label lblItmType 
         Caption         =   "SALECODE"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblItmCode 
         Caption         =   "PAYCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lblVdrCode 
         Caption         =   "PRCCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblDspItmType 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblDspItmCode 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3960
         TabIndex        =   15
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblDspVdrCode 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3960
         TabIndex        =   14
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblItmName 
         Caption         =   "NEW KEY:"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label lblDocLine 
      Caption         =   "SALECODE"
      Height          =   240
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label lblDspDocLine 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4320
      TabIndex        =   22
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblDocNo 
      Caption         =   "SALECODE"
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lblDspDocNo 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1680
      TabIndex        =   20
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAPW0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private wsFormID As String
Private waScrItm As New XArrayDB

Private wcCombo As Control
Private wlItmID    As Long
Private wlVdrID    As Long
Private wsJCurr    As String
Private wdJExcr    As Double


Private wsDocLine    As String

Private wsBaseCurCd    As String

'variable for new property
Private mlKeyID    As Long
Private wlSoPtjID    As Long

Private mbResult  As Boolean

Property Let KeyID(ByVal NewKeyID As Long)

   mlKeyID = NewKeyID

End Property


Property Get Result() As Boolean

   Result = mbResult
   
End Property


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

    
    Call Load_JobInfo
    txtUnitPrice.Text = Format("0", gsUprFmt)
    txtQty.Text = Format("0", gsQtyFmt)
    txtMU.Text = Format("1", gsAmtFmt)
    txtAmt.Text = Format("0", gsAmtFmt)
    txtNet.Text = Format("0", gsAmtFmt)
        
    wlSoPtjID = 0
    wlItmID = 0
    wlVdrID = 0

    
    tblCommon.Visible = False
    txtUnitPrice.Locked = True
    
    FocusMe cboDocLine
    

End Sub









Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()

If cmdSave = False Then Exit Sub

mbResult = True
Ini_Scr
cboDocLine.Text = wsDocLine
cboItmType.SetFocus


End Sub

Private Sub Form_Load()
 
    MousePointer = vbHourglass
        
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Scr
  
  
    MousePointer = vbDefault


End Sub



Private Sub Form_Unload(Cancel As Integer)
 

 Set waScrItm = Nothing

 
    
End Sub



Private Sub Ini_Form()

    Me.KeyPreview = True
    wsFormID = "APW0011"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    mbResult = False
   

End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblDocLine.Caption = Get_Caption(waScrItm, "DOCLINE")
    lblItmType.Caption = Get_Caption(waScrItm, "ITMTYPE")
    lblItmCode.Caption = Get_Caption(waScrItm, "ITMCODE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblItmName.Caption = Get_Caption(waScrItm, "ITMNAME")
    lblQty.Caption = Get_Caption(waScrItm, "QTY")
    lblUnitPrice.Caption = Get_Caption(waScrItm, "UNITPRICE")
    lblMU.Caption = Get_Caption(waScrItm, "MU")
    lblAmt.Caption = Get_Caption(waScrItm, "AMT")
    lblNet.Caption = Get_Caption(waScrItm, "NET")
    
    
    btnOK.Caption = Get_Caption(waScrItm, "OK")
    btnCancel.Caption = Get_Caption(waScrItm, "CANCEL")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

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



Private Sub cboDocLine_GotFocus()
    FocusMe cboDocLine
End Sub

Private Sub cboDocLine_LostFocus()
    FocusMe cboDocLine, True
End Sub


Private Sub cboDocLine_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboDocLine, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboDocLine = False Then
                Exit Sub
        End If
        
        cboItmType.SetFocus
       
    End If
    
End Sub

Private Sub cboDocLine_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboDocLine
    
    wsSQL = "SELECT CONVERT(NVARCHAR(5), REPLICATE('0',3 - LEN(LTRIM(CONVERT(NVARCHAR(3),SOPTJDOCLINE))))  + CONVERT(NVARCHAR(3),SOPTJDOCLINE)), SOPTJDESC1 "
    wsSQL = wsSQL & "FROM SOASOPTJ "
    wsSQL = wsSQL & "WHERE SOPTJDOCID = " & mlKeyID & " "
    wsSQL = wsSQL & "ORDER BY SOPTJDOCLINE "
    
    Call Ini_Combo(2, wsSQL, cboDocLine.Left, cboDocLine.Top + cboDocLine.Height, tblCommon, wsFormID, "TBLDOCLINE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboDocLine() As Boolean
Dim wsDesc As String

    Chk_cboDocLine = False
     
    If Trim(cboDocLine.Text) = "" Then
        gsMsg = "必需輸入物料行列!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocLine.SetFocus
        Exit Function
    End If
    
    
    If Chk_DocLine(wsDesc) = False Then
        gsMsg = "沒有此物料行列!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocLine.SetFocus
        lblDspDocLine = ""
       Exit Function
    End If
    
    lblDspDocLine = wsDesc
    
    Chk_cboDocLine = True
    
End Function




Private Sub cboItmType_GotFocus()
    FocusMe cboItmType
End Sub

Private Sub cboItmType_LostFocus()
    FocusMe cboItmType, True
End Sub


Private Sub cboItmType_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboItmType, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmType = False Then
                Exit Sub
        End If
        
        cboITMCODE.SetFocus
       
    End If
    
End Sub

Private Sub cboItmType_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmType
    
    wsSQL = "SELECT ITMTYPECODE, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC") & " FROM MSTITEMTYPE WHERE ITMTYPECODE LIKE '%" & IIf(cboItmType.SelLength > 0, "", Set_Quote(cboItmType.Text)) & "%' "
    wsSQL = wsSQL & "AND ITMTYPESTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY ITMTYPECODE "
    
    Call Ini_Combo(2, wsSQL, cboItmType.Left, cboItmType.Top + cboItmType.Height, tblCommon, wsFormID, "TBLITMTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboItmType() As Boolean
Dim wsDesc As String

    Chk_cboItmType = False
     
    If Trim(cboItmType.Text) = "" Then
        gsMsg = "必需輸入物料分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmType.SetFocus
        Exit Function
    End If
    
    
    If Chk_ItemType(cboItmType, wsDesc) = False Then
        gsMsg = "沒有此物料分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmType.SetFocus
        lblDspItmType = ""
       Exit Function
    End If
    
    lblDspItmType = wsDesc
    
    Chk_cboItmType = True
    
End Function


Private Sub cboItmCode_GotFocus()
    FocusMe cboITMCODE
End Sub

Private Sub cboItmCode_LostFocus()
    FocusMe cboITMCODE, True
End Sub


Private Sub cboItmCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    Dim wdPrice As Double
    
    
    Call chk_InpLen(cboITMCODE, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmCode(wsDesc, wdPrice) = False Then
                Exit Sub
        End If
        
        txtItmName = wsDesc
        wdPrice = get_ItemSalePrice(CStr(wlItmID), CStr(wlVdrID), wsJCurr, wdJExcr)
        txtUnitPrice = Format(wdPrice, gsUprFmt)
        
        If Trim(cboVdrCode) = "" Then
            cboVdrCode.SetFocus
        Else
            txtQty.SetFocus
        End If
        
        Call Cal_Total
       
    End If
    
End Sub

Private Sub cboItmCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboITMCODE
    
    wsSQL = "SELECT ITMCODE, ITMITMTYPECODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME, STR(ITMDEFAULTPRICE,13,2) FROM mstITEM "
    wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(cboITMCODE.Text) & "%' "
    wsSQL = wsSQL & " AND ITMITMTYPECODE =  '" & Set_Quote(cboItmType.Text) & "' "
    wsSQL = wsSQL & " ORDER BY ITMCODE "
    
    
    Call Ini_Combo(4, wsSQL, cboITMCODE.Left, cboITMCODE.Top + cboITMCODE.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboItmCode(ByRef OutName As String, ByRef outPrice As Double) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    
    
    If Trim(cboITMCODE.Text) = "" Then
        gsMsg = "沒有輸入物料號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_cboItmCode = False
        lblDspItmCode.Caption = ""
        Exit Function
    End If
    
    
    wsSQL = "SELECT ITMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME, ITMPVDRID, ITMDEFAULTPRICE, ITMCURR FROM MSTITEM "
    wsSQL = wsSQL & " WHERE ITMCODE = '" & Set_Quote(cboITMCODE.Text) & "' AND ITMITMTYPECODE = '" & Set_Quote(cboItmType.Text) & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wlItmID = ReadRs(rsRcd, "ITMID")
        wlVdrID = To_Value(ReadRs(rsRcd, "ITMPVDRID"))
        lblDspItmCode = ReadRs(rsRcd, "ITNAME")
        OutName = ReadRs(rsRcd, "ITNAME")
        wsCurr = ReadRs(rsRcd, "ITMCURR")
        outPrice = To_Value(ReadRs(rsRcd, "ITMDEFAULTPRICE"))
        
        If wsCurr <> wsJCurr Then
        
        If getExcRate(wsCurr, gsSystemDate, wsExcr, "") = False Then
            wsExcr = "0"
        End If
    
        If To_Value(wdJExcr) = 0 Then
            outPrice = outPrice * To_Value(wsExcr)
        Else
            outPrice = outPrice * To_Value(wsExcr) / To_Value(wdJExcr)
        End If
        
        End If
        
        
        
        Chk_cboItmCode = True
    Else
        wlItmID = 0
        wlVdrID = 0
        lblDspItmCode = ""
        wsCurr = wsBaseCurCd
        txtUnitPrice = Format("0", gsUprFmt)
        OutName = ""
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboITMCODE.SetFocus
        Chk_cboItmCode = False
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    If wlVdrID = 0 Then
    
    wsSQL = "SELECT VdrItemVdrID VID, VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVdrItem, MstVendor "
    wsSQL = wsSQL & " WHERE VdrItemItmID = " & wlItmID
    wsSQL = wsSQL & " AND VdrItemStatus = '1' "
    wsSQL = wsSQL & " AND VdrItemVdrID = VdrID "
    wsSQL = wsSQL & " Order By VdrItemCostl "
    
    Else
    
    wsSQL = "SELECT VdrID VID, VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrID = " & wlVdrID
    
    End If
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        
        wlVdrID = ReadRs(rsRcd, "VID")
        cboVdrCode = ReadRs(rsRcd, "VdrCode")
        lblDspVdrCode = ReadRs(rsRcd, "VdrName")

    Else
    
        wlVdrID = 0
        cboVdrCode = ""
        lblDspVdrCode = ""

    End If


    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function



Private Sub cboVdrCode_GotFocus()
    FocusMe cboVdrCode
End Sub

Private Sub cboVdrCode_LostFocus()
    FocusMe cboVdrCode, True
End Sub


Private Sub cboVdrCode_KeyPress(KeyAscii As Integer)
Dim wdPrice As Double
    
    Call chk_InpLen(cboVdrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrCode = False Then
                Exit Sub
        End If
        
        wdPrice = get_ItemSalePrice(CStr(wlItmID), CStr(wlVdrID), wsJCurr, wdJExcr)
        txtUnitPrice = Format(wdPrice, gsUprFmt)
        txtQty.SetFocus
        
        Call Cal_Total
       
    End If
    
End Sub

Private Sub cboVdrCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrCode
    
    wsSQL = "SELECT VDRCODE, VDRNAME FROM MSTVENDOR "
    wsSQL = wsSQL & " WHERE VDRSTATUS <> '2' AND VDRCODE LIKE '%" & Set_Quote(cboVdrCode.Text) & "%' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VDRCODE "
    
    
    Call Ini_Combo(2, wsSQL, cboVdrCode.Left, cboVdrCode.Top + cboVdrCode.Height, tblCommon, wsFormID, "TBLVDRCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function chk_cboVdrCode() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    
    If Trim(cboVdrCode.Text) = "" Then
        gsMsg = "沒有輸入供應商!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        chk_cboVdrCode = False
        lblDspVdrCode.Caption = ""
        Exit Function
    End If
    
    
    wsSQL = "SELECT VdrID, VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode = '" & Set_Quote(cboVdrCode.Text) & "'"
    
     
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wlVdrID = To_Value(ReadRs(rsRcd, "VdrID"))
        lblDspVdrCode = ReadRs(rsRcd, "VdrName")
        chk_cboVdrCode = True
        
    Else
        wlVdrID = 0
        lblDspVdrCode = ""
        gsMsg = "沒有此供應商!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboVdrCode.SetFocus
        chk_cboVdrCode = False
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
        
     
End Function


Private Sub txtItmName_GotFocus()
    FocusMe txtItmName
End Sub

Private Sub txtItmName_LostFocus()
    FocusMe txtItmName, True
End Sub


Private Sub txtItmName_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtItmName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
                
        txtQty.SetFocus
        
    End If
    
End Sub


Private Sub txtQty_GotFocus()

    FocusMe txtQty
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtQty.Text, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtQty Then
            
            txtMU.SetFocus
            
            Call Cal_Total
            
        End If
    End If

End Sub

Private Function Chk_txtQty() As Boolean
    
    Chk_txtQty = False
    
    
    If To_Value(txtQty.Text) <= 0 Then
        gsMsg = "Qty Must > 0 !"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtQty.SetFocus
        Exit Function
    End If
    
    txtQty.Text = Format(txtQty.Text, gsQtyFmt)
    
    Chk_txtQty = True
    
End Function


Private Sub txtQty_LostFocus()
FocusMe txtQty, True
End Sub


Private Sub txtMU_GotFocus()

    FocusMe txtMU
    
End Sub

Private Sub txtMU_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtMU.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtMU Then
            
            txtAmt.SetFocus
            
            Call Cal_Total
            
        End If
    End If

End Sub

Private Function Chk_txtMU() As Boolean
    
    Chk_txtMU = False
    
    
    If To_Value(txtMU.Text) < 0 Then
        gsMsg = "Qty Must > 0 !"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtMU.SetFocus
        Exit Function
    End If
    
    txtMU.Text = Format(txtMU.Text, gsAmtFmt)
    
    Chk_txtMU = True
    
End Function


Private Sub txtMU_LostFocus()
FocusMe txtMU, True
End Sub


Private Sub txtAmt_GotFocus()

    FocusMe txtAmt
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtAmt.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
            
            txtNet.SetFocus
            
    End If
    
End Sub
    

Private Sub txtAmt_LostFocus()
txtAmt.Text = Format(txtAmt.Text, gsAmtFmt)
FocusMe txtAmt, True
End Sub


Private Sub txtNet_GotFocus()

    FocusMe txtNet
    
End Sub

Private Sub txtNet_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtNet.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
            
            btnOK.SetFocus
            
    End If
    
End Sub
    

Private Sub txtNet_LostFocus()
txtNet.Text = Format(txtNet.Text, gsAmtFmt)
FocusMe txtNet, True
End Sub

Private Sub Load_JobInfo()
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT SOHDDOCNO, SOHDCURR, SOHDEXCR  "
    wsSQL = wsSQL & " FROM SOASOHD "
    wsSQL = wsSQL & " WHERE SOHDDOCID = " & To_Value(mlKeyID) & " "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wsJCurr = ReadRs(rsRcd, "SOHDCURR")
        wdJExcr = To_Value(ReadRs(rsRcd, "SOHDEXCR"))
        lblDspDocNo.Caption = ReadRs(rsRcd, "SOHDDOCNO")
    '    lblDspDocLine.Caption = Format(ReadRs(rsRcd, "SOPTJDOCLINE"), "000") & ": " & ReadRs(rsRcd, "SOPTJDESC1")

    Else
        wsJCurr = wsBaseCurCd
        wdJExcr = 1
        lblDspDocNo.Caption = ""
     '   lblDspDocLine.Caption = ""
    End If
    
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    
End Sub


Private Sub Cal_Total()
    
    txtAmt.Text = Format(To_Value(txtUnitPrice) * To_Value(txtQty), gsAmtFmt)
    txtNet.Text = Format(To_Value(txtUnitPrice) * To_Value(txtQty) / To_Value(txtMU), gsAmtFmt)
                 
        
    
End Sub


Private Function InputValidation() As Boolean
    
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    
    If Not Chk_cboDocLine Then Exit Function
    If Not Chk_cboItmType Then Exit Function
    If Not Chk_cboItmCode("", 0) Then Exit Function
    If Not chk_cboVdrCode Then Exit Function
    If Not Chk_txtQty Then Exit Function
    If Not Chk_txtMU Then Exit Function
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
Private Function cmdSave() As Boolean
    
    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wsUpdFlg As String
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Function
    End If
    
    
    gsMsg = "你是否要更新項目售價?"
    If MsgBox(gsMsg, vbYesNo, gsTitle) = vbYes Then
    wsUpdFlg = "Y"
    Else
    wsUpdFlg = "N"
    End If
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_APW0011A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wlSoPtjID)
    Call SetSPPara(adcmdSave, 2, wlItmID)
    Call SetSPPara(adcmdSave, 3, wlVdrID)
    Call SetSPPara(adcmdSave, 4, txtItmName)
    Call SetSPPara(adcmdSave, 5, txtUnitPrice)
    Call SetSPPara(adcmdSave, 6, txtQty)
    Call SetSPPara(adcmdSave, 7, txtMU)
    Call SetSPPara(adcmdSave, 8, txtAmt)
    Call SetSPPara(adcmdSave, 9, txtNet)
    Call SetSPPara(adcmdSave, 10, wsUpdFlg)
    
    adcmdSave.Execute
    
    
    cnCon.CommitTrans
    
    
   gsMsg = "物料已儲存!"
   MsgBox gsMsg, vbOKOnly, gsTitle
    
    
    Set adcmdSave = Nothing
    cmdSave = True
    wsDocLine = cboDocLine.Text
    
    
    
    MousePointer = vbDefault
    
    Exit Function
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Function


Private Function Chk_DocLine(ByRef OutName As String) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    
    
    If Trim(cboDocLine.Text) = "" Then
        gsMsg = "沒有輸入物料行列!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_DocLine = False
        wlSoPtjID = 0
        lblDspDocLine.Caption = ""
        Exit Function
    End If
    
    
    wsSQL = "SELECT SOPTJID, SOPTJDESC1 FROM SOASOPTJ "
    wsSQL = wsSQL & " WHERE SOPTJDOCLINE = " & To_Value(cboDocLine.Text) & " AND SOPTJDOCID = " & To_Value(mlKeyID) & " "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wlSoPtjID = ReadRs(rsRcd, "SOPTJID")
        OutName = ReadRs(rsRcd, "SOPTJDESC1")
        
        
        Chk_DocLine = True
    Else
        wlSoPtjID = 0
        OutName = ""
         Chk_DocLine = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function


