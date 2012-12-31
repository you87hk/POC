Attribute VB_Name = "modHH"
Public Function ExportToHHFile(ByVal inPath As String, ByVal InTrnCd As String, ByVal InDocID As String, ByVal inMode As Integer, ByVal wsWhs As String) As Boolean
    'inID - Document ID
    'inPath - File Path for file to export to
    'inMode - Mode: 1 - Overwrite, 2 - Append
    
    Dim adExport As New ADODB.Recordset
    Dim adUpdFlg As New ADODB.Command
    Dim sBuffer As String
    
    
    Dim wiDocID As Integer
    Dim wsDefWhs As String

    
    
    
    'wsDefWhs = Get_WorkStation_Info("WSWHSCODE")
    wsDefWhs = wsWhs
    
    ExportToHHFile = False
    
    Select Case InTrnCd
    Case "PO"
    
    wsSQL = "SELECT 'GRN' AS TYPE ,POHDDOCNO AS DOCNO, POHDREFNO AS JOBNO, ITMBINNO LOC, "
    wsSQL = wsSQL & " ITMCODE, (PODTQTY- PODTSCHQTY) * PODTCF AS QTY, 0 LINE "
    wsSQL = wsSQL & " FROM POPPOHD, POPPODT, MSTITEM "
    wsSQL = wsSQL & " WHERE POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND POHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND PODTQTY > PODTSCHQTY "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO, ITMCODE "
    
    Case "GR"
    
    wsSQL = "SELECT 'GRN' AS TYPE ,GRHDDOCNO AS DOCNO, POHDREFNO AS JOBNO, ITMBINNO LOC, "
    wsSQL = wsSQL & " ITMCODE, GRDTQTY * GRDTCF AS QTY, 0 LINE "
    wsSQL = wsSQL & " FROM POPGRHD, POPGRDT, POPPOHD, MSTITEM "
    wsSQL = wsSQL & " WHERE GRHDDOCID = GRDTDOCID "
    wsSQL = wsSQL & " AND GRHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND GRHDREFDOCID = POHDDOCID "
    wsSQL = wsSQL & " AND GRDTITEMID = ITMID "
    wsSQL = wsSQL & " ORDER BY GRHDDOCNO, ITMCODE "
    
    
    Case "SP"
    
     wsSQL = "SELECT 'PCK' AS TYPE ,SOHDDOCNO AS DOCNO, SOHDDOCNO AS JOBNO, SPDTWHSCODE LOC, "
    wsSQL = wsSQL & " SOPTJDOCLINE LINE, ITMCODE, SUM(SPDTQTY - SPDTOUTQTY) QTY "
    wsSQL = wsSQL & " From SOASOHD, SOASPHD, SOASPDT, MSTITEM, SOASODT, SOASOPTJ "
    wsSQL = wsSQL & " Where SPHDDOCID = SPDTDOCID "
    wsSQL = wsSQL & " AND SPHDREFDOCID = SOHDDOCID "
    wsSQL = wsSQL & " AND SOHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND SPDTITEMID = ITMID "
    wsSQL = wsSQL & " AND SPDTQTY > SPDTOUTQTY "
    wsSQL = wsSQL & " AND SPDTSODTID = SODTID "
    wsSQL = wsSQL & " AND SPDTSOID = SODTDOCID "
    wsSQL = wsSQL & " AND SODTPTJID  = SOPTJID "
    wsSQL = wsSQL & " GROUP BY SOHDDOCNO, SOPTJDOCLINE, SPDTWHSCODE, ITMCODE "
    
    Case "SW"
    
     wsSQL = "SELECT 'PCK' AS TYPE ,SOHDDOCNO AS DOCNO, SOHDDOCNO AS JOBNO, SWDTWHSCODE LOC, "
    wsSQL = wsSQL & " SOPTJDOCLINE LINE, ITMCODE, SUM(SWDTQTY - SWDTOUTQTY) QTY "
    wsSQL = wsSQL & " From SOASOHD, SOASWHD, SOASWDT, MSTITEM, SOASODT, SOASOPTJ "
    wsSQL = wsSQL & " Where SWHDDOCID = SWDTDOCID "
    wsSQL = wsSQL & " AND SWHDREFDOCID = SOHDDOCID "
    wsSQL = wsSQL & " AND SOHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND SWDTITEMID = ITMID "
    wsSQL = wsSQL & " AND SWDTQTY > SWDTOUTQTY "
    wsSQL = wsSQL & " AND SWDTSODTID = SODTID "
    wsSQL = wsSQL & " AND SWDTSOID = SODTDOCID "
    wsSQL = wsSQL & " AND SODTPTJID  = SOPTJID "
    wsSQL = wsSQL & " GROUP BY SOHDDOCNO, SOPTJDOCLINE, SWDTWHSCODE, ITMCODE "
    
    Case "SO"
    
    wsSQL = "SELECT 'PCK' AS TYPE ,SOHDDOCNO AS DOCNO, SOHDDOCNO AS JOBNO, SODTWHSCODE LOC, "
    wsSQL = wsSQL & " SOPTJDOCLINE LINE, ITMCODE, SUM(SODTTOTQTY - SODTSCHQTY) QTY "
    wsSQL = wsSQL & " FROM SOASOHD, SOASOPTJ, SOASODT, MSTITEM "
    wsSQL = wsSQL & " WHERE SOHDDOCID = SOPTJDOCID "
    wsSQL = wsSQL & " AND SOPTJID = SODTPTJID "
    wsSQL = wsSQL & " AND SOHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
    wsSQL = wsSQL & " AND SODTTOTQTY > SODTSCHQTY "
    wsSQL = wsSQL & " GROUP BY SOHDDOCNO, SODTWHSCODE, SOPTJDOCLINE, ITMCODE "
    
    
    Case "TR"
    
    wsSQL = "SELECT CASE WHEN SJDTTRNTYPE = 'STKIN' THEN 'TRI' ELSE 'TRO' END AS TYPE ,SJHDDOCNO AS DOCNO, SJHDJOBNO AS JOBNO, SJDTWHSCODE LOC, "
    wsSQL = wsSQL & " ITMCODE, SJDTQTY QTY, 0 LINE "
    wsSQL = wsSQL & " FROM ICSTKADJ, ICSTKADJDT, MSTITEM "
    wsSQL = wsSQL & " WHERE SJHDDOCID = SJDTDOCID "
    wsSQL = wsSQL & " AND SJHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND SJDTITEMID = ITMID "
    wsSQL = wsSQL & " AND SJDTWHSCODE = '" & wsDefWhs & "' "
    
    
    Case "DN"
    
    wsSQL = "SELECT 'DN' AS TYPE ,SOHDDOCNO AS DOCNO, SOHDDOCNO AS JOBNO, SODTWHSCODE LOC, "
    wsSQL = wsSQL & " SOPTJDOCLINE LINE, ITMCODE, SUM(SODTTOTQTY) QTY "
    wsSQL = wsSQL & " FROM SOASOHD, SOASOPTJ, SOASODT, MSTITEM "
    wsSQL = wsSQL & " WHERE SOHDDOCID = SOPTJDOCID "
    wsSQL = wsSQL & " AND SOPTJID = SODTPTJID "
    wsSQL = wsSQL & " AND SOHDDOCID = " & To_Value(InDocID) & " "
    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
    wsSQL = wsSQL & " GROUP BY SOHDDOCNO, SODTWHSCODE, SOPTJDOCLINE, ITMCODE "
    
    
    End Select
    

    adExport.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If adExport.RecordCount <= 0 Then
        gsMsg = "No Record can be匯出!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        adExport.Close
        Set adExport = Nothing
        Exit Function
    End If
    
    adExport.MoveFirst
    
    If inMode = 1 Then
        Open inPath For Output As #1
    Else
        Open inPath For Append As #1
    End If
        Do While Not adExport.EOF
            
        
            sBuffer = Set_FixLenStr(ReadRs(adExport, "TYPE"), 3)
            sBuffer = sBuffer & Set_FixLenStr(ReadRs(adExport, "DOCNO"), 15)
            sBuffer = sBuffer & Set_FixLenStr(ReadRs(adExport, "JOBNO"), 15)
            sBuffer = sBuffer & Set_FixLenStr(ReadRs(adExport, "LOC"), 10)
            sBuffer = sBuffer & Set_FixLenStr(ReadRs(adExport, "ITMCODE"), 30)
            sBuffer = sBuffer & Set_FixLenInt(0, 8)     'HHT Qty
          '  sBuffer = sBuffer & Set_FixLenInt(To_Value(ReadRs(adExport, "QTY")), 8)
            sBuffer = sBuffer & Set_FixLenLong(To_Value(ReadRs(adExport, "QTY")), 8)
          
            sBuffer = sBuffer & Set_FixLenStr("", 1)    'Match Flag
            sBuffer = sBuffer & Set_FixLenStr("", 10)   'Staff ID
            sBuffer = sBuffer & Set_FixLineNo(ReadRs(adExport, "JOBNO"), To_Value(ReadRs(adExport, "LINE")))
            sBuffer = sBuffer & Set_FixLenStr("", 1)    'ABC Flag
            
            
            'Write #1, sBuffer
            Print #1, sBuffer
            
            
            adExport.MoveNext
        Loop
    Close #1
    
    adExport.Close
    Set adExport = Nothing
    
    wsGenDte = gsSystemDate
    
  '  cnCon.BeginTrans
  '  wsSql = "UPDATE soaLXHD SET LXHDUPDFLG = 'Y' "
  '  wsSql = wsSql & " WHERE LXHDDOCID = " & wiDocID

   ' cnCon.Execute wsSql
   '    cnCon.CommitTrans

    ExportToHHFile = True
End Function


Public Function ImportFromHH(ByVal InUserID As String, ByVal inDteTim As String, ByVal inPath As String) As Boolean
    'inUserID - User ID
    'inDteTim - Date time
    'inPath - File Path for file to import from
    
Dim sfile As String
Dim sFileName As String
Dim sExt As String


Dim SDATE As String
Dim iNum As Integer
Dim sType As String

Dim sBuffer As String

Dim wsTrnCd As String
Dim wsDocNo As String
Dim wsJobNo As String
Dim wsLoc As String
Dim wsItem As String
Dim wiHHQty As String
Dim wiQty As String
Dim wsMatch As String
Dim wsFirst As String
Dim wsStaff As String
Dim wsLine As String
Dim wsABC As String


Dim adCmd As New ADODB.Command

ImportFromHH = False

On Error GoTo ImportFromHH_Err




sfile = Right(inPath, Len(inPath) - InStrRev(inPath, "\"))
sFileName = Left(sfile, InStr(sfile, ".") - 1)
sExt = Right(sfile, Len(sfile) - InStr(sfile, "."))
wsFirst = "Y"

If UCase(sExt) = "TXT" Or UCase(sExt) = "XLS" Or UCase(sExt) = "DOC" Or UCase(sExt) = "BAK" Or UCase(sExt) = "FLD" Then
   Exit Function
End If

SDATE = Year(Now) & "/" & Left(sFileName, 2) & "/" & Mid(sFileName, 3, 2)
iNum = To_Value(Mid(sFileName, 5, 2))
sType = Right(sFileName, 1)

cnCon.BeginTrans
Set adCmd.ActiveConnection = cnCon


    Open inPath For Input As #1
        Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, sBuffer  ' Read data into two variables.
            
            Select Case sType
            
            Case "D"
            
            wsTrnCd = Left(sBuffer, 3)
            wsDocNo = Mid(sBuffer, 4, 15)
            wsJobNo = Mid(sBuffer, 19, 15)
            wsLoc = Mid(sBuffer, 34, 10)
            wsItem = Mid(sBuffer, 44, 30)
            wiHHQty = To_Value(Mid(sBuffer, 74, 8))
            wiQty = To_Value(Mid(sBuffer, 82, 8))
            wsMatch = Mid(sBuffer, 90, 1)
            
            wsStaff = Mid(sBuffer, 91, 10)
            wsLine = To_Value(Mid(sBuffer, 101, 3))
            wsABC = To_Value(Mid(sBuffer, 104, 1))
            
            
            adCmd.CommandText = "USP_HHIM001"
            adCmd.CommandType = adCmdStoredProc
            adCmd.Parameters.Refresh
            
            Call SetSPPara(adCmd, 1, InUserID)
            Call SetSPPara(adCmd, 2, inDteTim)
            Call SetSPPara(adCmd, 3, sFileName & sExt)
            Call SetSPPara(adCmd, 4, sExt)
            Call SetSPPara(adCmd, 5, SDATE)
            Call SetSPPara(adCmd, 6, iNum)
            
            Call SetSPPara(adCmd, 7, wsTrnCd)
            Call SetSPPara(adCmd, 8, Trim(wsDocNo))
            Call SetSPPara(adCmd, 9, wsJobNo)
            Call SetSPPara(adCmd, 10, wsLoc)
            Call SetSPPara(adCmd, 11, Trim(wsItem))
            Call SetSPPara(adCmd, 12, wiHHQty)
            Call SetSPPara(adCmd, 13, wiQty)
            Call SetSPPara(adCmd, 14, wsMatch)
            Call SetSPPara(adCmd, 15, wsFirst)
            
            Call SetSPPara(adCmd, 16, wsStaff)
            Call SetSPPara(adCmd, 17, wsLine)
            Call SetSPPara(adCmd, 18, wsABC)
            
            adCmd.Execute
            
            wsFirst = "N"
            
            
            Case "S"
            
            wsTrnCd = "STK"
            wsDocNo = ""
            wsJobNo = ""
            wsLoc = Left(sBuffer, 10)
            wsItem = Mid(sBuffer, 11, 30)
            wiHHQty = To_Value(Mid(sBuffer, 41, 8))
            wiQty = 0
            wsMatch = ""
            
            wsStaff = ""
            wsLine = 0
            wsABC = ""
            
            
            adCmd.CommandText = "USP_HHIM001"
            adCmd.CommandType = adCmdStoredProc
            adCmd.Parameters.Refresh
            
            Call SetSPPara(adCmd, 1, InUserID)
            Call SetSPPara(adCmd, 2, inDteTim)
            Call SetSPPara(adCmd, 3, sFileName & sExt)
            Call SetSPPara(adCmd, 4, sExt)
            Call SetSPPara(adCmd, 5, SDATE)
            Call SetSPPara(adCmd, 6, iNum)
            
            Call SetSPPara(adCmd, 7, wsTrnCd)
            Call SetSPPara(adCmd, 8, Trim(wsDocNo))
            Call SetSPPara(adCmd, 9, wsJobNo)
            Call SetSPPara(adCmd, 10, wsLoc)
            Call SetSPPara(adCmd, 11, Trim(wsItem))
            Call SetSPPara(adCmd, 12, wiHHQty)
            Call SetSPPara(adCmd, 13, wiQty)
            Call SetSPPara(adCmd, 14, wsMatch)
            Call SetSPPara(adCmd, 15, wsFirst)
            
            Call SetSPPara(adCmd, 16, wsStaff)
            Call SetSPPara(adCmd, 17, wsLine)
            Call SetSPPara(adCmd, 18, wsABC)
            
            adCmd.Execute
            
            wsFirst = "N"
            
            
           Case "U"
            
            wsTrnCd = "USR"
            wsDocNo = ""
            wsJobNo = ""
            wsLoc = ""
            wsItem = ""
            wiHHQty = 0
            wiQty = 0
            wsMatch = ""
            wsStaff = Left(sBuffer, 10)
            wsLine = 0
            wsABC = ""
            
            
            adCmd.CommandText = "USP_HHIM001"
            adCmd.CommandType = adCmdStoredProc
            adCmd.Parameters.Refresh
            
            Call SetSPPara(adCmd, 1, InUserID)
            Call SetSPPara(adCmd, 2, inDteTim)
            Call SetSPPara(adCmd, 3, sFileName & sExt)
            Call SetSPPara(adCmd, 4, sExt)
            Call SetSPPara(adCmd, 5, SDATE)
            Call SetSPPara(adCmd, 6, iNum)
            
            Call SetSPPara(adCmd, 7, wsTrnCd)
            Call SetSPPara(adCmd, 8, Trim(wsDocNo))
            Call SetSPPara(adCmd, 9, wsJobNo)
            Call SetSPPara(adCmd, 10, wsLoc)
            Call SetSPPara(adCmd, 11, Trim(wsItem))
            Call SetSPPara(adCmd, 12, wiHHQty)
            Call SetSPPara(adCmd, 13, wiQty)
            Call SetSPPara(adCmd, 14, wsMatch)
            Call SetSPPara(adCmd, 15, wsFirst)
            
            Call SetSPPara(adCmd, 16, wsStaff)
            Call SetSPPara(adCmd, 17, wsLine)
            Call SetSPPara(adCmd, 18, wsABC)
            
            adCmd.Execute
            
            wsFirst = "N"
            
            End Select
        Loop
    Close #1

cnCon.CommitTrans
Set adCmd = Nothing

'If wsFirst = "Y" Then
'    gsMsg = sFileName & " 沒有紀錄匯入!"
'    MsgBox gsMsg, vbOKOnly, gsTitle
'End If


ImportFromHH = True
Exit Function

ImportFromHH_Err:
        
        
    MsgBox "ImportFromHH_Err!" & Err.Description
    cnCon.RollbackTrans
    
    Set adCmd = Nothing

End Function


Public Function Chk_HHNo(ByVal InHHNo As String, ByRef OutUpdFlg As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    

    wsSQL = "SELECT HHUPDFLG "
    wsSQL = wsSQL & "FROM SYSHHIM001 "
    wsSQL = wsSQL & "WHERE HHNO = '" & Set_Quote(InHHNo) & "' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutUpdFlg = ReadRs(rsRcd, "HHUPDFLG")
        Chk_HHNo = True
    Else
        OutUpdFlg = ""
        Chk_HHNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function
Public Function SendToHH(ByVal inPath As String) As Boolean

    Dim X As Long
    Dim DosCmd As String
    
    

On Error GoTo SendToHH_Err
    
    
    DosCmd = gsHHPath & "\IT3CW32.EXE +B115200 +P1 " & inPath & " +E +V"
    X = Shell(DosCmd, vbMaximizedFocus)
    
    SendToHH = True
    
    Exit Function
    
SendToHH_Err:

        
    MsgBox "SendToHH_Err!" & Err.Description
    SendToHH = False

End Function

Public Function ReceiveFromHH(ByVal inPath As String) As Boolean

    Dim X As Long
    Dim DosCmd As String
    
    
On Error GoTo ReceiveFromHH_Err
    
    
    DosCmd = gsHHPath & "\IT3CW32.EXE +RC +B115200 +P1 +V +E +L0 " & inPath & "(FILE)"
    
    X = Shell(DosCmd, vbMaximizedFocus)
    
    ReceiveFromHH = True
    
    
    Exit Function
    
ReceiveFromHH_Err:

  MsgBox "ReceiveFromHH_Err!" & Err.Description
    ReceiveFromHH = False

    

End Function


Public Function Init_Excel_BC(woExcelSheet1 As Object, _
                    xlApp2 As Object) As Integer

   Dim xlSheet2 As Object

' Variables used to format the selection criteria
' within the MS Excel Header and Footer
   Dim wsLeftHdr As String
   Dim wsCenterHdr As String
   Dim wsRightHdr As String
   Dim wsLeftFtr As String
   Dim wsRightFtr As String
   Dim wsTxt As String
   Dim wsLeftTxt As String
   Dim wsMid As String
   Dim wiLenHdr As Integer
   Dim wiLenFtr As Integer
   Dim wiFor As Integer
   Dim wiCtr As Integer
   Dim wiOldCtr As Integer
   Dim wbOk As Boolean
   Dim wbDot As Boolean
   Dim wbTxtEmpty As Boolean
   Dim wsTitle As String
   
' End of variables

   Init_Excel_BC = False

   'GET THE REQUIRED DATA FROM RPTHDR
   On Error GoTo Excel_Err
   'woExcelSheet1.PageSetup.PrintTitleRows = woExcelSheet1.Rows(1).Address
    woExcelSheet1.PageSetup.PrintTitleRows = "$1:$2"
    woExcelSheet1.PageSetup.PrintTitleColumns = "$A:$A"

   ' ************************** WORKSHEET (DATA) ************************** '
   'PAGE SETUP
   With woExcelSheet1.PageSetup
      'CREATE FOOTER
      wsLeftFtr = "&""Times New Roman""&6"
      wsRightFtr = "&""Times New Roman""&8Page &P of &N"
      wiLenFtr = Len(wsLeftFtr) + Len(wsRightFtr)
      
      'CREATE HEADER
      wsRightHdr = "&""Times New Roman""&8USER: " & gsUserID & vbLf & "   DATE: " & Change_SQLDate(gsSystemDate)
      wsCenterHdr = "&""Times New Roman,Bold""&10" & wsTitle
      wsLeftHdr = "&""Times New Roman""&8" & wsTitle & "&6" & vbLf & vbLf
      wiLenHdr = Len(wsRightHdr) + Len(wsCenterHdr) + Len(wsLeftHdr)
      
        
      'SET HEADER
      .LeftHeader = wsLeftHdr
      .CenterHeader = wsCenterHdr
      .RightHeader = wsRightHdr
      
      'SET FOOTER
      .LeftFooter = wsLeftFtr & IIf(wbDot = True, String(10, "."), "")
      .CenterFooter = ""
      .RightFooter = wsRightFtr
      
      'SET ALIGNMENT
      .CenterHorizontally = True
      .CenterVertically = False
      .TopMargin = xlApp2.InchesToPoints(0.78)
      .BottomMargin = xlApp2.InchesToPoints(0.35)
      .LeftMargin = xlApp2.InchesToPoints(0.2)
      .RightMargin = xlApp2.InchesToPoints(0.2)
      .HeaderMargin = xlApp2.InchesToPoints(0.2)
      .FooterMargin = xlApp2.InchesToPoints(0.2)
      
      'SET PAGE ORIENTATION
      .Orientation = xlLandscape

   End With

   ''GENERAL CELL/S SETUP
   'With woExcelSheet1
   '   .Cells.Select
   '   .Cells.Font.FontStyle = "Regular"
   '   .Cells.Font.Name = "Times New Roman"
   '   .Cells.Font.Size = 8
   '   .Cells.Borders.Weight = xlThin
   '   .Range("A1").Select
   'End With


   Set xlSheet2 = Nothing
   Init_Excel_BC = True

   Exit Function

Excel_Err:
   MsgBox "Exel_err"
   On Error Resume Next
   Set xlSheet2 = Nothing
   
End Function



Public Sub LoadExcel_BC(ByVal inFilePath As String)

    Dim xlApp As Excel.Application
    
    On Error GoTo LoadExcel_BC_Err
    
        If UCase(inFilePath) Like "*" & UCase(Dir(inFilePath)) & "*" And _
            Trim(Dir(inFilePath)) <> "" Then
            MousePointer = vbHourglass
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Visible = False
            xlApp.Workbooks.Open (Trim$(inFilePath)), ReadOnly:=True
            xlApp.Visible = True
        Else
            gsMsg = "File Not Existing!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        End If
    
    Set xlApp = Nothing
    MousePointer = vbDefault
    
    Exit Sub
    
LoadExcel_BC_Err:
    MsgBox "Can't Display: LoadExcel_BC_err!"
    Set xlApp = Nothing
    
End Sub

Public Function PrintExcel_BC(ByVal InTrnCd As String, ByVal InDocID As String) As Boolean

   
   Dim xlApp      As Object
   Dim xlBook      As Object
   Dim xlSheet1   As Object
   Dim rsRcd   As New ADODB.Recordset
   Dim rsItmRcd   As New ADODB.Recordset
      
   Dim OldCusCd As String
   
   Dim wsSQL      As String   'SQL statement
   Dim wlCtr      As Long
   Dim NoOfRecord As Long     'Number of Record
   Dim TotRow     As Long
   Dim CurRow     As Long
   Dim CurCol     As Integer
   Dim NoOfLabal  As Integer
   Dim i As Integer
   
   
   Dim wsExcelName As String
   
      
   On Error GoTo Excel_Print_Err1
   
   'Load Rcd
   PrintExcel_BC = False
   
   
    Select Case InTrnCd
    Case "ITM"
   
    wsSQL = "SELECT ITMCODE CODE, ITMCHINAME INAME, SUM((PODTQTY- PODTSCHQTY) * PODTCF) AS QTY "
    wsSQL = wsSQL & " FROM POPPODT, MSTITEM "
    wsSQL = wsSQL & " WHERE PODTDOCID IN (" & InDocID & ") "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND PODTQTY > PODTSCHQTY "
    wsSQL = wsSQL & " GROUP BY ITMCODE, ITMCHINAME "
    wsSQL = wsSQL & " ORDER BY ITMCODE, ITMCHINAME "
    
    
    wsExcelName = gsHHPath & "barcode\ITMCODE.XLS"

    Case "JOB"
    
    wsSQL = "SELECT POHDREFNO CODE, '' INAME, SUM((PODTQTY- PODTSCHQTY) * PODTCF) AS QTY "
    wsSQL = wsSQL & " FROM POPPOHD, POPPODT, MSTITEM "
    wsSQL = wsSQL & " WHERE POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND PODTDOCID IN (" & InDocID & ") "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND PODTQTY > PODTSCHQTY "
    wsSQL = wsSQL & " AND ISNULL(POHDREFNO,'') <> '' "
    wsSQL = wsSQL & " GROUP BY POHDREFNO "
    wsSQL = wsSQL & " ORDER BY POHDREFNO "
    
    wsExcelName = gsHHPath & "barcode\JOBNO.XLS"
    
    End Select
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
       ' gsMsg = IIf(gsLangID = "1", "No Data at this period or selection!", "沒有數據在此選擇!")
       ' MsgBox gsMsg, vbOKOnly, gsTitle
        PrintExcel_BC = True
        Exit Function
    End If
    
    'CONSTRUCT THE SELECT FIELDS STRING AND SUMMARY STRING
    
   On Error GoTo Excel_Print_Err2
   
   Set xlApp = CreateObject("EXCEL.Application")
   xlApp.Visible = False
   xlApp.SheetsInNewWorkbook = 1
   Set xlBook = xlApp.Workbooks.Add
   'Set xlSheet1 = xlApp.Workbooks(1).Worksheets(1)
   Set xlSheet1 = xlBook.Sheets.Add
   
'   If Init_Excel_BC(xlSheet1, xlApp) = False Then GoTo Excel_Print_Err1
   
   
   
   'Construct the summary select statement
   NoOfRecord = rsRcd.RecordCount
   wlCtr = 0
   TotRow = 0
   CurCol = 1
   CurRow = 1
   
   
    rsRcd.MoveFirst
         Do While Not rsRcd.EOF
         
             NoOfLabal = To_Value(ReadRs(rsRcd, "QTY"))
             
             If CurCol >= 200 Then
                 CurCol = 2
                 wlCtr = 0
                 Set xlSheet1 = xlBook.Sheets.Add
            '     If Init_Excel_BC(xlSheet1, xlApp) = False Then GoTo Excel_Print_Err1
              End If
                
              If wlCtr = 0 Then
                xlSheet1.Name = InTrnCd
                xlSheet1.Cells(CurRow, 1) = IIf(InTrnCd = "ITM", "ITEM", "JOB NO")
                xlSheet1.Cells(CurRow, 1).Interior.ColorIndex = 15
                xlSheet1.Cells(CurRow, 2) = IIf(InTrnCd = "ITM", "DESCRIPTION", "")
                xlSheet1.Cells(CurRow, 2).Interior.ColorIndex = 15
                
                CurRow = CurRow + 1
              End If
                
              If NoOfLabal > 0 Then
               For i = 1 To NoOfLabal
                xlSheet1.Cells(CurRow, 1).Value = ReadRs(rsRcd, "CODE")
                xlSheet1.Cells(CurRow, 2).Value = ReadRs(rsRcd, "INAME")
                CurRow = CurRow + 1
               Next i
               End If
                
                
    '            With xlSheet1.Range(xlSheet1.Cells(1, 2).Address, xlSheet1.Cells(CurRow - 1, CurCol + 2).Address)
    '            .Interior.ColorIndex = 15
    '            .WrapText = True
    '            .Borders(xlTop).Weight = xlMedium
    '            .Borders(xlRight).Weight = xlMedium
    '            .Borders(xlLeft).Weight = xlMedium
    '            .Borders(xlBottom).Weight = xlMedium
                '.AutoFit
    '            End With
            
'            xlSheet1.Cells(CurRow, CurCol).NumberFormat = gsAmtFmt
'            xlSheet1.Cells(CurRow, CurCol + 2).Borders(xlRight).Weight = xlMedium
            
            
            
         wlCtr = wlCtr + 1
         TotRow = TotRow + 1
         rsRcd.MoveNext
         Loop
         
        rsRcd.Close
        Set rsRcd = Nothing
        
        
   
PrintExcel_BC_Save:

   
      xlApp.Workbooks(1).Worksheets(1).SaveAs wsExcelName
      xlApp.Quit
      Set xlSheet1 = Nothing
      Set xlApp = Nothing
      MsgBox wsExcelName & " Finish"
      PrintExcel_BC = True
      
'      Call LoadExcel(txtFilePath.Text)
 
   
   Exit Function
   
Excel_Print_Err1:
   MsgBox "Excel_Print_Err1"
   On Error Resume Next
   Set xlSheet1 = Nothing
   Set xlApp = Nothing
   
   Exit Function

Excel_Print_Err2:
   If Err.Number = 1004 Then
        PrintExcel_BC = True
   Else
        MsgBox Err.Number & " Excel_Print_Err2"
   End If
   On Error Resume Next
   xlApp.Workbooks("BOOK1").Saved = True
   xlApp.Workbooks.Close
   xlApp.Quit
   Set xlSheet1 = Nothing
   Set xlApp = Nothing

End Function
Public Function Set_FixLenStr(ByVal inText As String, ByVal inLen As Integer) As String
    Dim iCtr As Integer
    Dim outText As String
    
    outText = Trim(inText)
        
    If Len(inText) < inLen Then
        For iCtr = 1 To inLen - Len(outText)
                outText = outText + " "
            
        Next iCtr
    Else
        outText = Left$(outText, inLen)
    End If
    Set_FixLenStr = outText
        
End Function

Public Function Set_FixLenInt(ByVal inNum As Integer, ByVal inLen As Integer) As String
    Dim iCtr As Integer
    Dim outText As String
    
    
    
    outText = inNum
        
    If Len(outText) < inLen Then
        For iCtr = 1 To inLen - Len(outText)
               ' outText = " " + outText
                outText = outText + " "
  
       Next iCtr
    Else
        outText = Right(outText, inLen)
    End If
    Set_FixLenInt = outText
        
End Function

Public Function Set_FixLenLong(ByVal inNum As Long, ByVal inLen As Integer) As String
    Dim iCtr As Integer
    Dim outText As String
    
    
    
    outText = inNum
        
    If Len(outText) < inLen Then
        For iCtr = 1 To inLen - Len(outText)
               ' outText = " " + outText
                outText = outText + " "
  
       Next iCtr
    Else
        outText = Right(outText, inLen)
    End If
    Set_FixLenLong = outText
        
End Function


Public Function Set_FixLineNo(ByVal inJob As String, ByVal inNum As Integer) As String
    Dim iCtr As Integer
    Dim outText As String
    
    If Trim(inJob) = "" Then
        outText = Set_FixLenStr("", 3)
    Else
        If inNum = 0 Then
            outText = Set_FixLenInt(0, 3)
        Else
            outText = Format(inNum, "000")
        End If
    
    End If
    
    Set_FixLineNo = outText
        
End Function
