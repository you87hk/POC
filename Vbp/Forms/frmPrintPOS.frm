VERSION 5.00
Begin VB.Form frmPrintPOS 
   BorderStyle     =   1  '單線固定
   Caption         =   "Report Printing"
   ClientHeight    =   1140
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5130
   Icon            =   "frmPrintPOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5130
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   792
      Top             =   360
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      ScaleHeight     =   360
      ScaleWidth      =   4545
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   4605
   End
End
Attribute VB_Name = "frmPrintPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Declare the constant value
'Option
Private Const cmdPreview   As Integer = 0
Private Const cmdPrinter   As Integer = 1
Private Const cmdExcel     As Integer = 2
Private Const cmdListView  As Integer = 3
'Command button for printing
Private Const PrintOK      As Integer = 0
Private Const PrintCancel  As Integer = 1
Private Const PrintSave    As Integer = 2
Private Const PrintPrinter As Integer = 3
Private Const PrintDetail  As Integer = 4
Private Const PrintFont    As Integer = 5
'Tab control
Private Const TabSelect    As Integer = 0
Private Const TabSort      As Integer = 1
'Command button for listview
Private Const MoveRight    As Integer = 0
Private Const MoveRightAll As Integer = 1
Private Const MoveLeft     As Integer = 2
Private Const MoveLeftAll  As Integer = 3
Private Const MoveUp       As Integer = 4
Private Const MoveDown     As Integer = 5
'Listview for from or to
Private Const PrintFrom    As Integer = 0
Private Const PrintTo      As Integer = 1
'Listview field subitem
Private Const ItemField    As Integer = 1
Private Const ItemNumFlag  As Integer = 2

Private Const cmdStart     As Integer = 1
Private Const cmdProcess   As Integer = 2
Private Const cmdFail      As Integer = 3

Private Const SummaryHeight   As Long = 3000
Private Const DetailHeight    As Long = 8000
Private Const FormWidth       As Long = 9816


Private Const tcPreview = "Preview"
Private Const tcPrint = "Print"
Private Const tcExcel = "Excel"
Private Const tcBrowse = "Browse"
Private Const tcPrinter = "Printer"
Private Const tcDetail = "Detail"
Private Const tcExit = "Exit"

'variable for new property

Private msReportID   As String
Private msRptDteTim  As String
Private msRptTitle   As String
Private miModID As Integer

Dim wsServer      As String
Dim wsDatabase    As String
Dim wsUser        As String
Dim wsPassword    As String
Dim wsSavePath    As String
Dim wiStatus      As Integer
Dim wiNoOfCopy    As Long
Dim wsQuery       As String
Dim wiNoOfRecords As Long
Dim wsAction      As Integer
Dim waScrItm      As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wiStartFlg    As Boolean
Dim wsRptPath As String
Dim wsExcPath As String

    
Private wsMsg As String

'Dim gdbSTEPPro As New ADODB.Connection
Dim WithEvents wdbCon As ADODB.Connection
Attribute wdbCon.VB_VarHelpID = -1



Property Let RptTitle(ByVal NewRptTitle As String)

   msRptTitle = NewRptTitle
   
End Property



Property Let RptDteTim(ByVal NewRptDteTim As String)

   msRptDteTim = NewRptDteTim
   
End Property



Property Let ReportID(ByVal NewReportID As String)

   msReportID = NewReportID

End Property


Property Let ModID(ByVal NewModID As Integer)

   miModID = NewModID

End Property

Private Sub Form_Activate()

   Me.MousePointer = vbHourglass

   Ini_Scr
   Ini_Caption

   
   
   If gsRTAccess = "N" Then
   Access_Print miModID
   Else
   Access_RunTimePrint miModID
   End If
   
   Me.MousePointer = vbDefault
   Unload Me
   Exit Sub
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Me.MousePointer = vbHourglass
   Me.Visible = False
   Me.KeyPreview = True
   Me.MousePointer = vbDefault
   
End Sub


Private Sub Form_LostFocus()

   'Unload Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'SaveUserDefault
   
    Set frmPrintPOS = Nothing
   
End Sub



Private Sub Timer1_Timer()
   wiStatus = 100
   If wiStatus < 100 Then
      wiStatus = wiStatus + 1
      Timer1.Interval = Timer1.Interval + 200
   Else
      wiStatus = 0
      Timer1.Enabled = False
   End If
   UpdateStatus picStatus, wiStatus
    
End Sub


Private Sub Ini_Scr()
   

   
   Me.Visible = True
   


End Sub





Private Sub Ini_Caption()

    Me.Caption = msRptTitle
   
End Sub

Private Sub Access_Print(inMode As Integer)
    ' Declare Database variable and Microsoft Access Form variable.
    Dim wsPara As String
'    Dim appAccess As Access.Application
    Dim appAccess As Object
    Dim tmpString As String
    Dim accessHandle As Long

    
    On Error GoTo Access_Print_Err
    
    UpdateStatus picStatus, 10
    
    ' Return instance of Application object.
    MousePointer = vbHourglass
    
    UpdateStatus picStatus, 20
    Set appAccess = CreateObject("Access.Application")
    UpdateStatus picStatus, 30
    accessHandle = appAccess.hWndAccessApp
    UpdateStatus picStatus, 40
    

'    appAccess.OpenCurrentDatabase (gsrptpath)
    appAccess.OpenCurrentDatabase gsDBPath & gsDBName, False
     UpdateStatus picStatus, 50
    
    wsQuery = "RPTUSRID = '" & Set_Quote(gsUserID) & "' "
    wsQuery = wsQuery & "   AND RPTDTETIM = #" & Change_SQLDate(msRptDteTim) & "# "

    
    UpdateStatus picStatus, 60
    
    ' Print report from Microsoft Access
    
   
    
                  
    Select Case inMode      'Select Print Mode
        Case 0      'Preview
            appAccess.DoCmd.OpenReport msReportID, acPreview, , wsQuery
            UpdateStatus picStatus, 70
            appAccess.DoCmd.Maximize
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
            
            
            appAccess.Visible = True
            Call SetWindowPos(accessHandle, HWND_TOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
        
        Case 1      'Print only
            appAccess.DoCmd.OpenReport msReportID, acNormal, , wsQuery
            UpdateStatus picStatus, 70
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
        
    End Select
    
    MousePointer = vbDefault
    Set appAccess = Nothing
    
    Exit Sub
    
Access_Print_Err:
    Set appAccess = Nothing
    MsgBox "Error in Access Print!"
    Exit Sub
    
End Sub


Private Sub Access_RunTimePrint(inMode As Integer)
    ' Declare Database variable and Microsoft Access Form variable.
    Dim wsPara As String
'    Dim appAccess As Access.Application
    Dim appAccess As Object
    Dim tmpString As String
    Dim accessHandle As Long
    Dim X As Long
    
    On Error GoTo Access_RunTimePrint_Err
    
    UpdateStatus picStatus, 10
    
    ' Return instance of Application object.
    MousePointer = vbHourglass
    
    
    
    UpdateStatus picStatus, 20
    
    X = Shell(gsRTPath & "msaccess.exe " & _
    Chr$(34) & gsDBPath & gsDBName & Chr$(34) & _
    "/Runtime /Wrkgrp " & Chr$(34) & _
    gsRTPath & "system.mdw" & Chr$(34), vbMaximizedFocus)
    
    UpdateStatus picStatus, 30
    Set appAccess = GetObject(gsDBPath & gsDBName)
    UpdateStatus picStatus, 40
    accessHandle = appAccess.hWndAccessApp
    UpdateStatus picStatus, 50
    

    'appAccess.OpenCurrentDatabase wsRptPath & gsDBName, False
     'UpdateStatus picStatus, 50
    
    wsQuery = "RPTUSRID = '" & Set_Quote(gsUserID) & "' "
    wsQuery = wsQuery & "   AND RPTDTETIM = #" & Change_SQLDate(msRptDteTim) & "# "
    
    
    
 '   wsPara = "ODBC;DRIVER=SQL Server;SERVER=" & wsServer & ";UID=" & wsUser & ";PWD=" & wsPassword & ";DATABASE=" & wsDatabase & ";"
    
    UpdateStatus picStatus, 60
    
    ' Print report from Microsoft Access
  '  appAccess.Run "connect_report", Trim$(wsPara), Trim$(Me.TableID)
    
     
                  
    Select Case inMode      'Select Print Mode
        Case 0      'Preview
        
            appAccess.DoCmd.OpenReport msReportID, acPreview, , wsQuery
            UpdateStatus picStatus, 70
            appAccess.DoCmd.Maximize
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
            
            
          '  appAccess.Visible = True
            Call SetWindowPos(accessHandle, HWND_TOPMOST, 0, 0, Screen.Width, Screen.Height, SWP_SHOWWINDOW)
        
        Case 1      'Print only
        
            appAccess.DoCmd.OpenReport msReportID, acNormal, , wsQuery
            UpdateStatus picStatus, 70
            UpdateStatus picStatus, 80
            UpdateStatus picStatus, 90
            UpdateStatus picStatus, 100, True
        
    End Select
    
    
  '  Debug.Assert 1 = 3

  '  appAccess.Quit acExit
    Set appAccess = Nothing
    
    MousePointer = vbDefault
    
    Exit Sub
    
Access_RunTimePrint_Err:
    Set appAccess = Nothing
    MsgBox "Error in Access Print! " & Err.Description
    Exit Sub
    
End Sub







Private Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase)
'-----------------------------------------------------------
' SUB: UpdStatusBar
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
    Dim intPercent As Double
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    If IsMissing(fBorderCase) Then fBorderCase = False
    
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &HFF0000 ' blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    intPercent = sngPercent
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    If sngPercent <> 0 Then
       strPercent = Format$(intPercent) & "%"
    Else
       strPercent = ""
    End If
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
       pic.Line (0, 0)-(pic.Width * (sngPercent / 100), pic.Height), pic.ForeColor, BF
    Else
       pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If
    
    pic.Refresh

End Sub
