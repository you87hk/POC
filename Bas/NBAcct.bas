Attribute VB_Name = "NBAcct"

Public Function Chk_MLClass(ByVal inCode As String, ByVal inType As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT MLDesc FROM mstMerchClass WHERE MLCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And MLStatus = '1' "
    wsSQL = wsSQL & "And MLType = '" & inType & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "MLDesc")
        Chk_MLClass = True
        
    Else
    
        OutDesc = ""
        Chk_MLClass = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function
Public Function Get_DueDte(ByVal inPayCode As String, ByVal inDocDate As String) As String
    
    Dim adDueDte As New ADODB.Command
    
    On Error GoTo Get_DueDte_Err
    
    Set adDueDte.ActiveConnection = cnCon
        
    adDueDte.CommandText = "USP_PAYDUE"
    adDueDte.CommandType = adCmdStoredProc
    adDueDte.Parameters.Refresh
    
    Call SetSPPara(adDueDte, 1, inPayCode)
    Call SetSPPara(adDueDte, 2, Change_SQLDate(inDocDate))
    adDueDte.Execute
    
    Get_DueDte = GetSPPara(adDueDte, 3)
    
    Set adDueDte = Nothing
    
    Exit Function
    
Get_DueDte_Err:
    MsgBox "Get_DueDte Error!"
    Set adDueDte = Nothing
    
End Function

Public Function Get_FiscalPeriod(ByVal inDate) As String
    
    Dim adCommand As New ADODB.Command
    
    On Error GoTo Get_FiscalPeriod_Err
    
    Set adCommand.ActiveConnection = cnCon
        
    adCommand.CommandText = "USP_getFiscalPeriod"
    adCommand.CommandType = adCmdStoredProc
    adCommand.Parameters.Refresh
    
    Call SetSPPara(adCommand, 1, inDate)
    adCommand.Execute
    
    Get_FiscalPeriod = GetSPPara(adCommand, 2) & GetSPPara(adCommand, 3)
    
    Set adCommand = Nothing
    
    Exit Function
    
Get_FiscalPeriod_Err:
    MsgBox "Get_FiscalPeriod Error!"
    Set adCommand = Nothing
    
End Function

Public Function Chk_VouPrefix(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT * FROM sysVouNo WHERE VouPrefix = '" & Set_Quote(inCode) & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        Chk_VouPrefix = True
    Else
        Chk_VouPrefix = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function


Public Function Chk_AutoVou(ByVal inDocTyp As String) As String

    Dim rst As New ADODB.Recordset
    Dim Criteria As String
    
    Chk_AutoVou = "N"
    
    Criteria = "SELECT VouLastKey FROM sysVouNo "
    Criteria = Criteria & " WHERE VouPrefix = '" & Set_Quote(inDocTyp) & "'"
    rst.Open Criteria, cnCon, adOpenStatic, adLockOptimistic
    
    If rst.RecordCount > 0 Then
       Chk_AutoVou = ReadRs(rst, "VouLastKey")
    End If
    
    rst.Close
    Set rst = Nothing
    
End Function

Public Function Chk_CurrInMod(InCurr As String, inMod As String) As Boolean
    
    Dim rsCurCod As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_CurrInMod = False
    
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE, SYSMONCTL "
    wsSQL = wsSQL & " WHERE EXCMN = CONVERT(INTEGER, MCCTLMN) "
    wsSQL = wsSQL & " AND EXCYR = MCCTLYR "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & " AND MCMODNO = '" & inMod & "' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
     
    
    rsCurCod.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsCurCod.RecordCount <= 0 Then
        rsCurCod.Close
        Set rsCurCod = Nothing
        Exit Function
    End If
    
    Chk_CurrInMod = True
    
    rsCurCod.Close
    Set rsCurCod = Nothing
    
End Function
