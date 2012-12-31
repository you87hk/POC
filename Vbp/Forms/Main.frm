VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "新雅文貨事業有限公司"
   ClientHeight    =   7965
   ClientLeft      =   4110
   ClientTop       =   1905
   ClientWidth     =   7680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Customer"
            Object.ToolTipText     =   "客戶"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Book"
            Object.ToolTipText     =   "書本"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "QuotationEntry"
            Object.ToolTipText     =   "Quotation Entry"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sale Notes"
            Object.ToolTipText     =   "訂貨單"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Invoice"
            Object.ToolTipText     =   "Invoice"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "PurchaseOrder"
            Object.ToolTipText     =   "Purchase Order"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "DeliveryNotes"
            Object.ToolTipText     =   "Delivery Notes"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "DepositEntry"
            Object.ToolTipText     =   "Deposit Entry"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Customerpayment"
            Object.ToolTipText     =   "Customer Payment"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "BillPayment"
            Object.ToolTipText     =   "Bill Payment"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ShipIn"
            Object.ToolTipText     =   "Ship In"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "OrderAllocation"
            Object.ToolTipText     =   "Order Allocation"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ShipOut"
            Object.ToolTipText     =   "Ship Out"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ToDoList"
            Object.ToolTipText     =   "To Do List"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "RegisterInquiry"
            Object.ToolTipText     =   "Register Inquiry"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Reporting"
            Object.ToolTipText     =   "Reporting"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出系統"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList iglMain 
      Left            =   240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":223E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2690
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3286
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":35A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":38BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":402E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":434A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":479E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":522A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5546
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5862
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5B7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  '對齊表單下方
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   7575
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7858
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2000/10/12"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "PM 07:34"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "檔案"
      Begin VB.Menu mnuUser 
         Caption         =   "User Master"
      End
      Begin VB.Menu mnuKey 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "退出系統"
      End
   End
   Begin VB.Menu mnuMasterMenu 
      Caption         =   "主資料"
      Begin VB.Menu mnuCustomer 
         Caption         =   "客戶"
      End
      Begin VB.Menu mnuVendor 
         Caption         =   "供應商"
      End
      Begin VB.Menu mnuSalesman 
         Caption         =   "營業員"
      End
      Begin VB.Menu mnuNature 
         Caption         =   "性質"
      End
      Begin VB.Menu mnuMethod 
         Caption         =   "銷售渠道"
      End
      Begin VB.Menu mnuPayTerm 
         Caption         =   "付款條款"
      End
      Begin VB.Menu mnuTerritory 
         Caption         =   "地區"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "貨幣"
      End
      Begin VB.Menu mnuExchangeRate 
         Caption         =   "對換率"
      End
      Begin VB.Menu Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBook 
         Caption         =   "書本"
      End
      Begin VB.Menu mnuSaleDiscount 
         Caption         =   "銷售折扣"
      End
      Begin VB.Menu mnuCategory 
         Caption         =   "杜威分類"
      End
      Begin VB.Menu mnuCategoryDiscount 
         Caption         =   "圖書折扣類別"
      End
      Begin VB.Menu mnuItemType 
         Caption         =   "圖書分類"
      End
      Begin VB.Menu mnuAccountType 
         Caption         =   "會計版別"
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "語種"
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "程度"
      End
      Begin VB.Menu mnuPrintSize 
         Caption         =   "開度"
      End
      Begin VB.Menu mnuPackingType 
         Caption         =   "裝幀"
      End
      Begin VB.Menu mnuShip 
         Caption         =   "貨運"
      End
      Begin VB.Menu mnuUOM 
         Caption         =   "量度單位"
      End
      Begin VB.Menu mnuStore 
         Caption         =   "點存"
      End
      Begin VB.Menu mnuPriceTerm 
         Caption         =   "銷售條款"
      End
      Begin VB.Menu mnuWarehouse 
         Caption         =   "貨倉"
      End
      Begin VB.Menu mnuRemark 
         Caption         =   "註解"
      End
      Begin VB.Menu mnuMerchClass 
         Caption         =   "買手"
      End
      Begin VB.Menu Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemPrice 
         Caption         =   "書本價格對換"
      End
      Begin VB.Menu mnuCReg 
         Caption         =   "客戶登錄號"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "營業清單"
      Begin VB.Menu mnuSN 
         Caption         =   "訂貨單"
      End
      Begin VB.Menu mnuConvertSNToSC 
         Caption         =   "轉為銷售單"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSC 
         Caption         =   "銷售單"
      End
      Begin VB.Menu mnuExpSO 
         Caption         =   "匯出銷售單"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImpPO 
         Caption         =   "匯入採購單"
      End
      Begin VB.Menu mnuPO 
         Caption         =   "採購單"
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBAT001 
         Caption         =   "Batch Maintenance"
      End
   End
   Begin VB.Menu mnuCommunication 
      Caption         =   "資料轉送"
      Begin VB.Menu mnuShipIn 
         Caption         =   "匯入入貨"
      End
      Begin VB.Menu mnuShipOut 
         Caption         =   "匯入出貨"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "匯入存貨"
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReturnIn 
         Caption         =   "回貨"
      End
      Begin VB.Menu mnuReturnOut 
         Caption         =   "退貨"
      End
   End
   Begin VB.Menu mnuInquiry 
      Caption         =   "查詢"
      Begin VB.Menu mnuInqCustomer 
         Caption         =   "客戶"
      End
      Begin VB.Menu mnuInqVendor 
         Caption         =   "供應商"
      End
      Begin VB.Menu mnuInqBook 
         Caption         =   "書本"
      End
      Begin VB.Menu mnuInqBacth 
         Caption         =   "批次檔"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInqInventory 
         Caption         =   "存貨"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "報告"
      Begin VB.Menu mnuRptOSOrder 
         Caption         =   "未完成銷售單"
      End
      Begin VB.Menu mnuRptOSPO 
         Caption         =   "未完成採購單"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptSO 
         Caption         =   "銷售單"
      End
      Begin VB.Menu mnuRptPO 
         Caption         =   "採購單"
      End
      Begin VB.Menu Sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptLabel 
         Caption         =   "標籤"
      End
      Begin VB.Menu mnuBookLabel 
         Caption         =   "書本標籤"
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPickingList 
         Caption         =   "出貨單"
      End
      Begin VB.Menu Sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBC001 
         Caption         =   "Book List"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "視窗"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private waScrItm As New XArrayDB

Private Sub MDIForm_Load()
    Me.WindowState = 2
    'Me.Caption = GetFormName("frmMain")
    'mnuFile.Caption = GetMenuName("frmMain", "mnuFile")
'    tbrMain.Buttons.Item("Enquiry").ToolTipText = GetToolTipNew("frmMain", "Enquiry", "tbrMain")
    
    'If Not xLang(Me) Then
    '    MsgBox GetErrorMessage("E0001"), vbCritical + vbOKOnly, "Error"
    'End If
    
    'Call xMenu(Me)


  'Call IniForm
       
  

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sMsg As String
    Dim sSQL As String
On Error GoTo ErrHand

    sMsg = "Are you sure to exit this system?" & Chr(10) & Chr(10)
    sMsg = sMsg & "請問你是不是肯定退出這系統?"

    If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, gsTitle) = vbYes Then
        
 
        sSQL = "DUMP TRANSACTION SUNYADB WITH NO_LOG"
        cnCon.Execute sSQL
        
        
        Unload frmMain
        
        
    Else
        Cancel = True
    End If

Exit Sub

ErrHand:
     MsgBox Err.Description
     Cancel = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'L000331
    Set waScrItm = Nothing
    End
End Sub


Private Sub mnuAccountType_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmAT001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuBAT001_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmBAT001
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuBC001_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmBC001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuBook_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmB001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuBookLabel_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmLB001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCategory_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmCAT001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCategoryDiscount_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmCD001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuConvertSCToPO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmEX001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuConvertSNToSC_Click()
    Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmCVT001
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCReg_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmCR001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCurrency_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmCUR001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCustomer_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmC001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuCVT001_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmCVT001
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuDelivery_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmDN000
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuDeposit_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmDEP000
    Me.MousePointer = vbNormal

End Sub

Private Sub mnuExchangeRate_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmEXC001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuExpSO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmEX001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuFExit_Click()
    Unload Me
End Sub

Private Sub mnuInvoice_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmINV000
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuImpPO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM004
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuInqBacth_Click()
   Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmINQ002
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuInqBook_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmINQ001
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuInqCustomer_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmINQ004
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuInqInventory_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmINQ003
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuInqVendor_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmINQ005
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuItemPrice_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIP001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuItemType_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIT001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub



Private Sub mnuKey_Click()
   Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmCHGPWD
    newForm.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuLanguage_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmL001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuLevel_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmLVL001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuMaster_Click()
 Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuMerchClass_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmML001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuMethod_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmM001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuNature_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmN001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuOrderAllocation_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmSTKALL0
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPackingType_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmPT001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPassword_Click()
    Me.MousePointer = vbHourglass
    'frmUSR001.Show vbModal
    Me.MousePointer = vbModal
End Sub

Private Sub mnuPassword1_Click()
    Me.MousePointer = vbHourglass
    'frmPasswordInput.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPayTerm_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmPYT001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPRIINQ_Click()
    Me.MousePointer = vbHourglass
    'frmPRIINQ.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPickingList_Click()
    Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmDN001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPO_Click()
    Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmPO001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPriceTerm_Click()
    Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmPR001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPrintSize_Click()
    Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmPS001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPurchaseOrder_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmPO000
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuPurTar_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmPURTAR
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuQuotation_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmQTN000
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRegistorInuiry_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmINQ000
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuReporting_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmRPT000
    Me.MousePointer = vbNormal

End Sub

Private Sub mnuRemark_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmRmk001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuReturnIn_Click()
     Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM005
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuReturnOut_Click()
 Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM006
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRptLabel_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmLB002
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRptOSOrder_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmOS001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRptOSPO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmOS002
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRptPO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmPO002
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuRptSO_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSO002
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuSaleDiscount_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSD001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuSalesman_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSLM001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuSaleTar_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmSALTAR
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuSC_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSO001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuShip_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSHP001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuSN_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSN001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuShipIn_Click()
 Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM002
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuShipOut_Click()
 Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmIM003
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuStockCount_Click()
    Me.MousePointer = vbHourglass
    'frmSTKCNT.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuStockInquiry_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmINQ001
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuStockTransfer_Click()
    Me.MousePointer = vbHourglass

    'LoadForm frmST000
    Me.MousePointer = vbNormal
End Sub


Private Sub mnuSN001_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmSN001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuStore_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmS001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuTerritory_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmTerr001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuToDoList_Click()
    Me.MousePointer = vbHourglass
    'LoadForm frmDOL000
    Me.MousePointer = vbNormal

End Sub

Private Sub mnuUOM_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmUOM001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuUser_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmUSR001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuVendor_Click()
    Dim newForm As Form
 
    Me.MousePointer = vbHourglass
    Set newForm = New frmV001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuWarehouse_Click()
Dim newForm As Form
    Me.MousePointer = vbHourglass
    Set newForm = New frmWH001
    LoadForm newForm
    Me.MousePointer = vbNormal
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Cardfile"
            'mnuCustomer_Click
        
        Case "Customer"
            mnuCustomer_Click
        
        Case "Book"
            'mnuProduct_Click
            mnuBook_Click
            
        Case "QuotationEntry"
            'mnuQuotation_Click
            
        Case "Sale Notes"
            mnuSN_Click
            
        Case "Invoice"
            'mnuInvoice_Click
            
        Case "PurchaseOrder"
            'mnuPurchaseOrder_Click
        
        Case "DeliveryNotes"
            'mnuDelivery_Click
        
        Case "DepositEntry"
            'mnuDeposit_Click
            
        Case "Customerpayment"
            'mnuCustomerPayment_Click
            
        Case "BillPayment"
            'mnuBill_Click
            
        Case "ShipIn"
            'mnuShipIn_Click
            
        Case "OrderAllocation"
            'mnuOrderAllocation_Click
        
        Case "ShipOut"
            'mnuShipOut_Click
            
        Case "ToDoList"
            'mnuToDoList_Click
            
        Case "RegisterInquiry"
            'mnuRegistorInuiry_Click
            
        Case "Reporting"
            'mnuReporting_Click
            
        Case "Exit"
            mnuFExit_Click
    End Select
End Sub

Private Sub LoadForm(f As Form)
   f.WindowState = 0
   f.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
   f.Show
   f.ZOrder 0
   
End Sub

Private Sub IniForm()
 '   Me.KeyPreview = True
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
End Sub
Private Sub Ini_Menu()
        
    ' First node with 'Root' as text.
    Call Get_Scr_Item("AR000", waScrItm)
        
    Me.Caption = Get_Caption(waScrItm, "SCRHDR")
    mnuEntry.Caption = Get_Caption(waScrItm, "ENTRY")
    mnuPrint.Caption = Get_Caption(waScrItm, "PRINT")
    mnuUpdating.Caption = Get_Caption(waScrItm, "UPDATING")
    mnuInquiry.Caption = Get_Caption(waScrItm, "INQUIRY")
    mnuReportFunction.Caption = Get_Caption(waScrItm, "REPORTFUNCTION")
    mnuOption.Caption = Get_Caption(waScrItm, "OPTION")
    mnuExt.Caption = Get_Caption(waScrItm, "EXIT")
    mnuHlp.Caption = Get_Caption(waScrItm, "HELP")
    mnuHlpCon.Caption = Get_Caption(waScrItm, "CONTENT")
    mnuHlpSrh.Caption = Get_Caption(waScrItm, "SEARCH")
    mnuHlpAbt.Caption = Get_Caption(waScrItm, "ABOUT")
    
    
    Call Ini_PgmMenu(mnuEntItm, 1, "AR000", "MNU", waEntItm)
    Call Ini_PgmMenu(mnuPrtItm, 2, "AR000", "MNU", waPrtItm)
    Call Ini_PgmMenu(mnuUpdItm, 5, "AR000", "MNU", waUpdItm)
    Call Ini_PgmMenu(mnuInqItm, 4, "AR000", "MNU", waInqItm)
    Call Ini_PgmMenu(mnuReportFunctionItm, 3, "AR000", "MNU", waReportFunctionItm)
                
    sbStatusBar.Panels(1).Text = gsComNam
End Sub
