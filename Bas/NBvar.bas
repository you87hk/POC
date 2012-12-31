Attribute VB_Name = "NBVar"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function SetForegroundWindow Lib "user32" _
     (ByVal hwnd As Long) As Long
Declare Function IsIconic Lib "user32" _
     (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" _
     (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long



#Const DEBUG_FLAG = True

Public giCurrIndex As Integer


Public gsMODULE   As String
Public gsExcPath  As String
Public gsRptPath  As String
Public gsComNam   As String
Public gsDteFmt   As String
Public gsHostName As String
Public gsHostLogin As String

'Global Minimum Date and Maximum Date
Public gsDateFrom  As String
Public gsDateTo As String


Public cnCon As New ADODB.Connection        '-- ADO connection for manual use.

Public gsLangID As String
Public gsTitle As String
Public gsConnectString As String     '-- Connection string.
Public gsUserID As String
Public gsSystemDate As String
Public gsWorkStationID As String
Public gsCompID As String
Public gsMsg As String
Public gsDBName As String
Public gsWhsCode As String
Public gsRTAccess As String
Public gsRTPath As String
Public gsHHPath As String



Global Const DefaultPage = -1
Global Const AddRec = 1
Global Const CorRec = 2
Global Const DelRec = 3
Global Const RevRec = 4
Global Const CorRO = 5


Global Const gsQtyFmt = "#,##0"
Global Const gsAmtFmt = "#,##0.00"
Global Const gsUprFmt = "#,##0.0000"
Global Const gsExrFmt = "#,##0.000000"

Global Const giAmtDp = 2
Global Const giUprDp = 4
Global Const giExrDp = 6
Global Const giQtyDp = 0


Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40


'Global Minimum Value and Maximum Value
Global Const gsMinVal = "-9999999999999.99"
Global Const gsMaxVal = "9999999999999.99"

'Global Const giTimeOut = 0

