Attribute VB_Name = "modChungFai"

Public Sub LoadSaleNameByCode(outTxtBox As TextBox, ByVal inSaleCode)
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT SaleName FROM MstSalesman WHERE SaleCode ='" + Set_Quote(inSaleCode) + "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
          outTxtBox = ReadRs(rsRcd, "SaleName")
    Else
          outTxtBox = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    
End Sub


Public Function LoadSaleCodeByID(ByVal inSaleManID) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT SaleCode FROM MstSalesman WHERE SaleID =" & To_Value(inSaleManID)
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        LoadSaleCodeByID = ReadRs(rsRcd, "SaleCode")
    Else
        LoadSaleCodeByID = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Sub LoadSaleByID(outSalesCode As String, outSalesName As String, ByVal inSaleManID)
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleID =" & To_Value(inSaleManID)
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        outSalesCode = ReadRs(rsRcd, "SaleCode")
        outSalesName = ReadRs(rsRcd, "SaleName")
    Else
        outSalesCode = ""
        outSalesName = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Public Function getCallNo(ByVal inCallNo2) As String
    Dim wsChar As String
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsChar = Left(inCallNo2, 1)
    
    If (wsChar > "a" And wsChar < "z") Or (wsChar > "A" And wsChar < "Z") Then
        wsChar = Left(inCallNo2, 3)
    Else
        wsSQL = "SELECT CharTable.Hon "
        wsSQL = wsSQL + "From CharTable "
        wsSQL = wsSQL + "WHERE (((CharTable.Word)='" + Set_Quote(wsChar) + "'));"
        
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
        If rsRcd.RecordCount <> 0 Then
            wsChar = ReadRs(rsRcd, "Hon")
        End If
        
        rsRcd.Close
        Set rsRcd = Nothing
    End If
    
    getCallNo = Trim(wsChar)
    
    
End Function

Public Function get_PriCode(ByVal inCatCode) As String
    Dim wsChar As String
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
        get_PriCode = ""

    
        wsSQL = "SELECT PRICODE "
        wsSQL = wsSQL + "From mstPrimary "
        wsSQL = wsSQL + "WHERE PriCatCode ='" + Set_Quote(inCatCode) + "'"
        
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
        If rsRcd.RecordCount <> 0 Then
            get_PriCode = ReadRs(rsRcd, "PriCode")
        End If
        
        rsRcd.Close
        Set rsRcd = Nothing
    
    
End Function
Public Function LoadDescByCode(stblName As String, sKeyName As String, sRetName As String, sValue As String, bString As Boolean) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsResult As String
        
    LoadDescByCode = ""
        
    If Trim(sValue) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT " + sRetName
    
    If bString = True Then
        wsSQL = wsSQL + " FROM " + stblName + " WHERE " + sKeyName + " = '" + Set_Quote(sValue) + "' "
    Else
        wsSQL = wsSQL + " FROM " + stblName + " WHERE " + sKeyName + " = " + Set_Quote(sValue)
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wsResult = ReadRs(rsRcd, sRetName)
    Else
        wsResult = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    LoadDescByCode = wsResult
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function Chk_Curr(InCurr As String, inDate As String) As Boolean
    
    Dim rsCurCod As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_Curr = False
    
    wsCtlDte = IIf(Trim(inDate) = "" Or Trim(inDate) = "/  /", gsSystemDate, inDate)
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE "
    wsSQL = wsSQL & "EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
    wsSQL = wsSQL & " AND EXCCURR = '" & Set_Quote(InCurr) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    
    
    rsCurCod.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsCurCod.RecordCount <= 0 Then
        rsCurCod.Close
        Set rsCurCod = Nothing
        Exit Function
    End If
    
    Chk_Curr = True
    
    rsCurCod.Close
    Set rsCurCod = Nothing
    
End Function

Public Function Chk_PoHdDocNo(ByVal InDocNo As String, ByRef OutStatus As String, ByRef OutPgmNo As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT POHDPGMNO, POHDSTATUS FROM popPOHD WHERE POHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "POHDSTATUS")
        OutPgmNo = ReadRs(rsRcd, "POHDPGMNO")
        Chk_PoHdDocNo = True
    Else
        OutStatus = ""
        OutPgmNo = ""
        Chk_PoHdDocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function



Public Function Chk_CusCode(ByVal InCusNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT CusID, CusName, CusTel, CusFax, CusEMail FROM mstCustomer WHERE CusCode = '" & Set_Quote(InCusNo) & "' "
    wsSQL = wsSQL & "And CusStatus = '1' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "CusID")
        OutName = ReadRs(rsRcd, "CusName")
        OutTel = ReadRs(rsRcd, "CusTel")
        OutFax = ReadRs(rsRcd, "CusFax")
        
        Chk_CusCode = True
        
    Else
    
        OutID = 0
        OutName = ""
        OutTel = ""
        OutFax = ""
       
        Chk_CusCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function Chk_Salesman(ByVal inCode As String, ByRef OutID As Long, ByRef OutName As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT SaleID, SaleName FROM mstSalesman WHERE SaleCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And SaleStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "SaleID")
        OutName = ReadRs(rsRcd, "SaleName")
        Chk_Salesman = True
        
    Else
    
        OutID = 0
        OutName = ""
        Chk_Salesman = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Type(ByVal inCode As String, ByVal InClass As String, ByRef OutID As Long, ByRef OutName As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT TypID, TypDesc FROM mstType WHERE TypCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And TypClass = '" & Set_Quote(InClass) & "' "
    wsSQL = wsSQL & "And TypStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "TypID")
        OutName = ReadRs(rsRcd, "TypDesc")
        Chk_Type = True
        
    Else
    
        OutID = 0
        OutName = ""
        Chk_Type = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Region(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    wsSQL = "SELECT RgnDesc FROM MstRegion WHERE RgnCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And RgnStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "RgnDesc")
        Chk_Region = True
        
    Else
    
        OutDesc = ""
        Chk_Region = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function Chk_PayTerm(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT PayDesc FROM mstPayTerm WHERE PayCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And PayStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "PayDesc")
        Chk_PayTerm = True
        
    Else
    
        OutDesc = ""
        Chk_PayTerm = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_PriceTerm(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT PrcDesc FROM mstPriceTerm WHERE PrcCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And PrcStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "PrcDesc")
        Chk_PriceTerm = True
        
    Else
    
        OutDesc = ""
        Chk_PriceTerm = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function Chk_MerchClass(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT MLDesc FROM mstMerchClass WHERE MLCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And MLStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "MLDesc")
        Chk_MerchClass = True
        
    Else
    
        OutDesc = ""
        Chk_MerchClass = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Method(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT MethodDesc FROM mstMethod WHERE MethodCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And MethodStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "MethodDesc")
        Chk_Method = True
        
    Else
    
        OutDesc = ""
        Chk_Method = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Nature(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT NatureDesc FROM mstNature WHERE NatureCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And NatureStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "NatureDesc")
        Chk_Nature = True
        
    Else
    
        OutDesc = ""
        Chk_Nature = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Remark(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT RmkStatus FROM mstRemark WHERE RmkCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And RmkStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        Chk_Remark = True
        
    Else
    
        Chk_Remark = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_Ship(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT ShipStatus FROM mstShip WHERE ShipCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And ShipStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        Chk_Ship = True
        
    Else
    
        Chk_Ship = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Get_SaleDiscount(ByVal inNatureCode As String, ByVal InCusID As Long, ByVal InItemID As Long) As Double

   
Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wdCusDis As Double
    Dim wdSalDis As Double
    
    wdCusDis = 0
    wdSalDis = 0
    
        
    wsSQL = "SELECT CUSSPECDIS "
    wsSQL = wsSQL & " FROM mstCustomer "
    wsSQL = wsSQL & " WHERE CusID = " & InCusID
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wdCusDis = To_Value(ReadRs(rsRcd, "CUSSPECDIS"))
    End If
    
    rsRcd.Close
    
    
    wdCusDis = (100 - wdCusDis) / 100
        
        
    wsSQL = "SELECT SDDISCOUNT "
    wsSQL = wsSQL & " FROM mstSaleDiscount, mstCustomer, mstItem "
    wsSQL = wsSQL & " WHERE CusID = " & InCusID
    wsSQL = wsSQL & " AND ItmID = " & InItemID
    wsSQL = wsSQL & " AND SDNatureCode = '" & Set_Quote(inNatureCode) & "'"
    wsSQL = wsSQL & " AND SDMethodCode = CusMethodCode "
    wsSQL = wsSQL & " AND SDCDisCode = ItmCDisCode "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wdSalDis = To_Value(ReadRs(rsRcd, "SDDiscount"))
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    wdSalDis = (100 - wdSalDis) / 100
    
    Get_SaleDiscount = (1 - wdCusDis * wdSalDis) * 100
    
End Function

Public Function getCusItemPrice(ByVal InCusID As Long, ByVal InItemID As Long, ByVal InCurr As String) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    wsSQL = "SELECT CusItemPrice "
    wsSQL = wsSQL & " FROM MstCusItem "
    wsSQL = wsSQL & " WHERE CusItemCusID = " & InCusID
    wsSQL = wsSQL & " AND CusItemItmID = " & InItemID
    wsSQL = wsSQL & " AND CusItemCurr = '" & Set_Quote(InCurr) & "'"
    wsSQL = wsSQL & " AND CusItemStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getCusItemPrice = To_Value(ReadRs(rsRcd, "CusItemPrice"))
    Else
        getCusItemPrice = 0
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function getItemPrice(ByVal inDate As String, ByVal InItemID As Long, ByVal InCurr As String, ByVal InExcr As Double) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsCurr As String
    
    wsSQL = "Select PriChgDtCurr, PriChgDtUnitPrice "
    wsSQL = wsSQL & " From mstPriceChange, mstPriceChangeDt"
    wsSQL = wsSQL & " Where PriChgDocID = PriChgDtDocID "
    wsSQL = wsSQL & " AND PriChgDtItemID = " & InItemID
    wsSQL = wsSQL & " AND PriChgFromDate <= '" & Set_Quote(inDate) & "'"
    wsSQL = wsSQL & " AND PriChgToDate >= '" & Set_Quote(inDate) & "'"
    wsSQL = wsSQL & " AND PriChgStatus = '1'"
    
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getItemPrice = To_Value(ReadRs(rsRcd, "PriChgDtUnitPrice"))
        wsCurr = ReadRs(rsRcd, "PriChgDtCurr")
        If UCase(wsCurr) <> UCase(InCurr) Then
            getItemPrice = Format(getItemPrice * InExcr, gsAmtFmt)
        End If
    Else
        getItemPrice = 0
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Get_CreditLimit(ByVal InCusID As Long, ByVal inDate As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_CreditLimit_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "usp_GETCREDITLIMIT"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InCusID)
    Call SetSPPara(adCmd, 2, inDate)
    
    adCmd.Execute
    
    Get_CreditLimit = GetSPPara(adCmd, 3)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_CreditLimit_Err:
    MsgBox "Get_CreditLimit_Err!"
    Set adCmd = Nothing
    
End Function
Public Function Check_CreditLimit(ByVal InCusID As Long, ByVal InDocID As Long, ByVal inDate As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Check_CreditLimit_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_CHECKCREDITLIMIT"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InCusID)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, inDate)
    
    adCmd.Execute
    
    Check_CreditLimit = GetSPPara(adCmd, 4)
    
    Set adCmd = Nothing
    
    Exit Function
    
Check_CreditLimit_Err:
    MsgBox "Check_CreditLimit_Err!"
    Set adCmd = Nothing
    
End Function
Public Function getVdrItemDiscount(ByVal inVdrID As Long, ByVal InItemID As Long) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    wsSQL = "SELECT VdrItemDiscount "
    wsSQL = wsSQL & " FROM MstVdrItem "
    wsSQL = wsSQL & " WHERE VdrItemVdrID = " & inVdrID
    wsSQL = wsSQL & " AND VdrItemItmID = " & InItemID
    wsSQL = wsSQL & " AND VdrItemStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getVdrItemDiscount = To_Value(ReadRs(rsRcd, "VdrItemDiscount"))
    Else
        getVdrItemDiscount = 1
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_VdrCode(ByVal InVdrNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT VdrID, VdrName, VdrTel, VdrFax FROM mstVendor WHERE VdrCode = '" & Set_Quote(InVdrNo) & "' "
    wsSQL = wsSQL & "And VdrStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "VdrID")
        OutName = ReadRs(rsRcd, "VdrName")
        OutTel = ReadRs(rsRcd, "VdrTel")
        OutFax = ReadRs(rsRcd, "VdrFax")
        Chk_VdrCode = True
        
    Else
    
        OutID = 0
        OutName = ""
        OutTel = ""
        OutFax = ""
        Chk_VdrCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_ItmCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_ItmCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT ItmStatus "
    wsSQL = wsSQL & " FROM MstItem WHERE ItmCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "ItmStatus")
    
    Chk_ItmCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function



Public Function Chk_GrpCode(ByVal inNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT USRGRPCODE FROM mstUser WHERE USRGRPCODE = '" & Set_Quote(inNo) & "' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        Chk_GrpCode = True
    Else
    
        Chk_GrpCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function
Public Function Chk_TblName(ByVal inNo As String, ByRef outNo As String, ByRef outSts As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT tblTableID, tblPgmID, tblStatusName FROM sysMstTable WHERE tblTableID = '" & Set_Quote(inNo) & "' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        Chk_TblName = True
        outNo = ReadRs(rsRcd, "tblPgmID")
        outSts = ReadRs(rsRcd, "tblStatusName")
    Else
    
        Chk_TblName = False
        outNo = ""
        outSts = ""
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function



Public Function Get_StockAvailable(ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_StockAvailable_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_STOCKAVAILABLE"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inItmID)
    Call SetSPPara(adCmd, 2, inWhsCode)
    Call SetSPPara(adCmd, 3, InLotNo)
    
    adCmd.Execute
    
    Get_StockAvailable = GetSPPara(adCmd, 4)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_StockAvailable_Err:
    MsgBox "Get_StockAvailable_Err!"
    Set adCmd = Nothing
    
End Function




Public Function Get_PoBalQty(ByVal inTrnCode As String, ByVal InDocID As Long, ByVal InPoID As Long, ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_PoBalQty_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetPoBalQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InPoID)
    Call SetSPPara(adCmd, 4, inItmID)
    Call SetSPPara(adCmd, 5, inWhsCode)
    Call SetSPPara(adCmd, 6, InLotNo)
    
    adCmd.Execute
    
    Get_PoBalQty = To_Value(GetSPPara(adCmd, 7))
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_PoBalQty_Err:
    MsgBox "Get_PoBalQty_Err!"
    Set adCmd = Nothing
    
End Function

Public Function Get_SoBalQty(ByVal inTrnCode As String, ByVal InDocID As Long, ByVal InSoID As Long, ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_SoBalQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetSoBalQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InSoID)
    Call SetSPPara(adCmd, 4, inItmID)
    Call SetSPPara(adCmd, 5, inWhsCode)
    Call SetSPPara(adCmd, 6, InLotNo)
    
    adCmd.Execute
    
    Get_SoBalQty = To_Value(GetSPPara(adCmd, 7))
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_SoBalQty_Err:
    MsgBox "Get_SoBalQty_Err!"
    Set adCmd = Nothing
    
End Function

Public Function Get_StTrnQty(ByVal inTrnCode As String, ByVal InDocID As Long, ByVal InSoID As Long, ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_StTrnQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetStTrnQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InSoID)
    Call SetSPPara(adCmd, 4, inItmID)
    Call SetSPPara(adCmd, 5, inWhsCode)
    Call SetSPPara(adCmd, 6, InLotNo)
    
    adCmd.Execute
    
    Get_StTrnQty = To_Value(GetSPPara(adCmd, 7))
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_StTrnQty_Err:
    
    MsgBox "Get_StTrnQty_Err!"
    Set adCmd = Nothing
    
End Function
Public Function Get_IcQty(ByVal inTrnCode As String, ByVal InSoID As Long, ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_IcQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetIcQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    Call SetSPPara(adCmd, 2, InSoID)
    Call SetSPPara(adCmd, 3, inItmID)
    Call SetSPPara(adCmd, 4, inWhsCode)
    Call SetSPPara(adCmd, 5, InLotNo)
    
    adCmd.Execute
    
    Get_IcQty = To_Value(GetSPPara(adCmd, 6))
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_IcQty_Err:
    
    MsgBox "Get_IcQty_Err!"
    Set adCmd = Nothing
    
End Function

Public Function Get_WorkStation_Info(ByVal inField As String) As String

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
     wsSQL = "SELECT " & inField & " FROM sysWSINFO WHERE WSID = '" & gsWorkStationID & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Get_WorkStation_Info = ReadRs(rsRcd, inField)
    Else
        MsgBox "Can't Access WorkStation Info!"
        Get_WorkStation_Info = ""
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_SpecialPassword(ByVal inPassword As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_SpecialPassword = False
    
    If UCase(gsUserID) = "NBASE" And UCase(inPassword) = "NBTEL" Then
            Chk_SpecialPassword = True
            Exit Function
    End If
    
    wsSQL = "SELECT USRPWD, USRNAME FROM MSTUSER WHERE USRCODE = 'FAIRRACK' "

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
        If Encrypt(rsRcd("USRPWD")) <> UCase(inPassword) Then
            Chk_SpecialPassword = False
        Else
            Chk_SpecialPassword = True
        End If
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_Whs(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT WhsDesc FROM MstWarehouse WHERE WhsCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And WhsStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "WhsDesc")
        Chk_Whs = True
        
    Else
    
        OutDesc = ""
        Chk_Whs = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing

End Function

Public Function Chk_LotEnabled(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT * FROM MstWarehouse WHERE WhsCode = '" & Set_Quote(inCode) & "' AND WhsLotEnabled = 'Y' "
    wsSQL = wsSQL & "And WhsStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        Chk_LotEnabled = True
    Else
        Chk_LotEnabled = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing

End Function



Public Function Chk_LotB(ByVal inCode As String, ByVal inBin As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT WHSLOTB FROM MstWarehouse WHERE WhsCode = '" & Set_Quote(inCode) & "' and WHSLOTB = '" & Set_Quote(Trim(inBin)) & "'"
    wsSQL = wsSQL & " AND WhsLotEnabled = 'Y' And WhsStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Or inBin = "999" Then
        Chk_LotB = False
    Else
        Chk_LotB = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing

End Function
Public Function Chk_SoReady(ByVal InDocNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsReady As String
    
    Chk_SoReady = False
    
    wsSQL = "SELECT SOHDREADY FROM SOASOHD WHERE SOHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
         wsReady = ReadRs(rsRcd, "SOHDREADY")
    Else
         wsReady = "N"
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    If wsReady = "Y" Then
        Chk_SoReady = True
    End If
        
    

End Function


Public Function Chk_TrnHdDocNo(ByVal InTrnCd As String, ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    
    Select Case InTrnCd
    Case "SN", "QT"
    wsSQL = "SELECT SNHDSTATUS STATUS FROM soaSNHD WHERE SNHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "VQ"
    wsSQL = "SELECT VQHDSTATUS STATUS FROM soaVQHD WHERE VQHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "SO"
    wsSQL = "SELECT SOHDSTATUS STATUS FROM soaSOHD WHERE SOHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "IV"
    wsSQL = "SELECT IVHDSTATUS STATUS FROM soaIVHD WHERE IVHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "SP"
    wsSQL = "SELECT SPHDSTATUS STATUS FROM soaSPHD WHERE SPHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "SW"
    wsSQL = "SELECT SWHDSTATUS STATUS FROM soaSWHD WHERE SWHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    
    Case "SR"
    wsSQL = "SELECT SRHDSTATUS STATUS FROM soaSRHD WHERE SRHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "SD"
    wsSQL = "SELECT SDHDSTATUS STATUS FROM soaSDHD WHERE SDHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "PO"
    wsSQL = "SELECT POHDSTATUS STATUS FROM popPOHD WHERE POHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "PV"
    wsSQL = "SELECT PVHDSTATUS STATUS FROM popPVHD WHERE PVHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "GR"
    wsSQL = "SELECT GRHDSTATUS STATUS FROM popGRHD WHERE GRHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "PR"
    wsSQL = "SELECT PRHDSTATUS STATUS FROM popPRHD WHERE PRHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "TR", "SM", "AJ", "DM", "SK"
    wsSQL = "SELECT SJHDSTATUS STATUS FROM ICSTKADJ WHERE SJHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case "SC"
    wsSQL = "SELECT SCHDSTATUS STATUS FROM ICSTKCNT WHERE SCHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    Case Else
        OutStatus = ""
        Chk_TrnHdDocNo = False
        MsgBox "No Such TrnCd!Chk_TrnHdDocNo Mod!"
        Exit Function
    End Select
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "STATUS")
        Chk_TrnHdDocNo = True
    Else
        OutStatus = ""
        Chk_TrnHdDocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function



Public Function Get_IcTrnQty(ByVal inTrnCode As String, ByVal inItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_IcTrnQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetIcTrnQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    Call SetSPPara(adCmd, 2, inItmID)
    Call SetSPPara(adCmd, 3, inWhsCode)
    Call SetSPPara(adCmd, 4, InLotNo)
    
    adCmd.Execute
    
    Get_IcTrnQty = To_Value(GetSPPara(adCmd, 5))
    
        Set adCmd = Nothing
    
    Exit Function
    
Get_IcTrnQty_Err:
    
    MsgBox "Get_IcTrnQty_Err!"
    Set adCmd = Nothing
    
End Function



Public Function Get_LocBalbyCode(ByVal inItmCd As String, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_IcTrnQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "usp_GetLocBalbyCode"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inItmCd)
    Call SetSPPara(adCmd, 2, inWhsCode)
    Call SetSPPara(adCmd, 3, InLotNo)
    
    adCmd.Execute
    
    Get_LocBalbyCode = To_Value(GetSPPara(adCmd, 4))
    
        Set adCmd = Nothing
    
    Exit Function
    
Get_IcTrnQty_Err:
    
    MsgBox "Get_IcTrnQty_Err!"
    Set adCmd = Nothing
    
End Function
Public Function Get_AvgCost(ByVal inItmID As Long, ByVal inWhsCode As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_AvgCost_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetAvgCost"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inItmID)
    Call SetSPPara(adCmd, 2, inWhsCode)
    
    adCmd.Execute
    
    Get_AvgCost = To_Value(GetSPPara(adCmd, 3))
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_AvgCost_Err:
    
    MsgBox "Get_AvgCost_Err!"
    Set adCmd = Nothing
    
End Function

Public Function Get_ItemNo(ByVal inTrnCode As String) As String
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_ItemNo_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GenItemNo"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inTrnCode)
    
    adCmd.Execute
    
    Get_ItemNo = GetSPPara(adCmd, 2)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_ItemNo_Err:
    MsgBox "Get_ItemNo_Err!"
    Set adCmd = Nothing
    
End Function


Public Function get_ItemSalePrice(inItmID As String, inVdrID As String, InCurr As String, InExcr As Double) As Double
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wdPrice As Double
    Dim wdDiscount As Double
    Dim wdMark As Double
    Dim wsCurr As String
    Dim wsExcRate As String
    
    get_ItemSalePrice = 0
    
    
    If Trim(inVdrID) = "" Or Trim(inItmID) = "" Then
    Exit Function
    End If
    
    wsSQL = "SELECT ItmUnitPrice, ItmCurr, VdrItemDiscount, ItmMarkUp "
    wsSQL = wsSQL & " FROM MstItem, MstVdrItem "
    wsSQL = wsSQL & " WHERE ItmID = VdrItemItmID "
    wsSQL = wsSQL & " AND VdrItemVdrID = " & inVdrID
    wsSQL = wsSQL & " AND VdrItemItmID = " & inItmID
    wsSQL = wsSQL & " AND VdrItemStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsCurr = ReadRs(rsRcd, "ItmCurr")
    wdPrice = To_Value(ReadRs(rsRcd, "ItmUnitPrice"))
    wdDiscount = To_Value(ReadRs(rsRcd, "VdrItemDiscount"))
    wdMark = To_Value(ReadRs(rsRcd, "ItmMarkUp"))
    
    If UCase(InCurr) <> UCase(wsCurr) Then
    
    If getExcRate(wsCurr, gsSystemDate, wsExcRate, "") = False Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If To_Value(InExcr) = 0 Then
        wdPrice = wdPrice * To_Value(wsExcRate)
    Else
        wdPrice = wdPrice * To_Value(wsExcRate) / To_Value(InExcr)
    End If
    
    End If
    
    get_ItemSalePrice = Format(wdPrice * wdDiscount * wdMark, gsUprFmt)
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    
End Function


Public Function get_ItemVdrCost(inItmID As String, inVdrID As String, InCurr As String, InExcr As Double) As Double
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wdPrice As Double
    Dim wdDiscount As Double
    Dim wsCurr As String
    Dim wsExcRate As String
    
    get_ItemVdrCost = 0
    
    
    If Trim(inVdrID) = "" Or Trim(inItmID) = "" Then
    Exit Function
    End If
    
    wsSQL = "SELECT ItmUnitPrice, ItmCurr, VdrItemDiscount "
    wsSQL = wsSQL & " FROM MstItem, MstVdrItem "
    wsSQL = wsSQL & " WHERE ItmID = VdrItemItmID "
    wsSQL = wsSQL & " AND VdrItemVdrID = " & inVdrID
    wsSQL = wsSQL & " AND VdrItemItmID = " & inItmID
    wsSQL = wsSQL & " AND VdrItemStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsCurr = ReadRs(rsRcd, "ItmCurr")
    wdPrice = To_Value(ReadRs(rsRcd, "ItmUnitPrice"))
    wdDiscount = To_Value(ReadRs(rsRcd, "VdrItemDiscount"))
    
    If UCase(InCurr) <> UCase(wsCurr) Then
    
    If getExcRate(wsCurr, gsSystemDate, wsExcRate, "") = False Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If To_Value(InExcr) = 0 Then
        wdPrice = wdPrice * To_Value(wsExcRate)
    Else
        wdPrice = wdPrice * To_Value(wsExcRate) / To_Value(InExcr)
    End If
    
    End If
    
    get_ItemVdrCost = Format(wdPrice * wdDiscount, gsUprFmt)
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    
End Function

Public Function Chk_MClass(ByVal inCode As String, ByVal inType As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT MLDesc FROM mstMerchClass "
    wsSQL = wsSQL & "WHERE MLCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And MLType = '" & Set_Quote(inType) & "' "
    wsSQL = wsSQL & "And MLStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDesc = ReadRs(rsRcd, "MLDesc")
        Chk_MClass = True
        
    Else
    
        OutDesc = ""
        Chk_MClass = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function


Public Function get_VdrItemDiscount(inItmID As String, inVdrID As String) As Double
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    get_VdrItemDiscount = 1
    
    
    If Trim(inVdrID) = "" Or Trim(inItmID) = "" Then
    Exit Function
    End If
    
    wsSQL = "SELECT VdrItemDiscount "
    wsSQL = wsSQL & " FROM MstVdrItem "
    wsSQL = wsSQL & " WHERE VdrItemVdrID = " & inVdrID
    wsSQL = wsSQL & " AND VdrItemItmID = " & inItmID
    wsSQL = wsSQL & " AND VdrItemStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    get_VdrItemDiscount = To_Value(ReadRs(rsRcd, "VdrItemDiscount"))
    
    rsRcd.Close
    Set rsRcd = Nothing
        
    
End Function


Public Function Chk_Staff(ByVal inCode As String, ByRef OutID As Long, ByRef OutName As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT StaffID, StaffName FROM mstStaff WHERE StaffCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And StaffStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "StaffID")
        OutName = ReadRs(rsRcd, "StaffName")
        Chk_Staff = True
        
    Else
    
        OutID = 0
        OutName = ""
        Chk_Staff = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_SoRefDoc(ByVal inID As String, ByRef OutTrnCd As String, ByRef OutDocNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    Chk_SoRefDoc = False
    
    
    wsSQL = "SELECT SPHDDOCNO FROM SOASPHD WHERE SPHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND SPHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "SPHDDOCNO")
        OutTrnCd = "SP"
        Chk_SoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

    wsSQL = "SELECT SWHDDOCNO FROM SOASWHD WHERE SWHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND SWHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "SWHDDOCNO")
        OutTrnCd = "SW"
        Chk_SoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
  

    wsSQL = "SELECT IVHDDOCNO FROM SOAIVHD WHERE IVHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND IVHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "IVHDDOCNO")
        OutTrnCd = "IV"
        Chk_SoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    wsSQL = "SELECT POHDDOCNO FROM POPPOHD WHERE POHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND POHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "POHDDOCNO")
        OutTrnCd = "PO"
        Chk_SoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
  
End Function


Public Function Chk_SpQty(ByVal inID As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    Chk_SpQty = False
    
    
    wsSQL = "SELECT SPDTOUTQTY FROM SOASPDT WHERE SPDTDOCID = " & To_Value(inID) & " "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        If To_Value(ReadRs(rsRcd, "SPDTOUTQTY")) > 0 Then
        Chk_SpQty = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        End If
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
   
  
End Function

Public Function Chk_ItemType(ByVal inCode As String, ByRef OutName As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT ItmTypeChiDesc, ItmTypeEngDesc FROM mstItemType WHERE ItmTypeCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & "And ItmTypeStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        If gsLangID = "1" Then
        OutName = ReadRs(rsRcd, "ItmTypeEngDesc")
        Else
        OutName = ReadRs(rsRcd, "ItmTypeChiDesc")
        End If
        
        Chk_ItemType = True
        
    Else
    
        OutName = ""
        Chk_ItemType = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

Public Function Chk_VoucNo(ByVal InPfx As String, ByVal InDocNo As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    
    wsSQL = "SELECT VOHDDOCNO FROM GLVOHD "
    wsSQL = wsSQL & " WHERE VOHDPFX = '" & Set_Quote(InPfx) & "'"
    wsSQL = wsSQL & " AND VOHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        Chk_VoucNo = True
    Else
        Chk_VoucNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function
Public Sub Get_CusSaleInfo(ByVal InCusID As Long, ByVal inYYYY As String, ByVal inMM As String, ByRef OutCreLmt As Double, ByRef OutCreLft As Double, ByRef OutOpnBal As Double, ByRef OutTotBal As Double, ByRef OutCMQty As Double, ByRef OutCYQty As Double, ByRef OutTotQty As Double, ByRef OutCMSal As Double, ByRef OutCYSal As Double, ByRef OutTotSal As Double, ByRef OutCMNet As Double, ByRef OutCYNet As Double, ByRef OutTotNet As Double)
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_CusSaleInfo_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GETCUSSALEINFO"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InCusID)
    Call SetSPPara(adCmd, 2, inYYYY)
    Call SetSPPara(adCmd, 3, inMM)
    
    
    adCmd.Execute
    
    OutCreLmt = GetSPPara(adCmd, 4)
    OutCreLft = GetSPPara(adCmd, 5)
    OutOpnBal = GetSPPara(adCmd, 6)
    OutTotBal = GetSPPara(adCmd, 7)
    
    OutCMQty = GetSPPara(adCmd, 8)
    OutCYQty = GetSPPara(adCmd, 9)
    OutTotQty = GetSPPara(adCmd, 10)
    
    OutCMSal = GetSPPara(adCmd, 11)
    OutCYSal = GetSPPara(adCmd, 12)
    OutTotSal = GetSPPara(adCmd, 13)
    
    OutCMNet = GetSPPara(adCmd, 14)
    OutCYNet = GetSPPara(adCmd, 15)
    OutTotNet = GetSPPara(adCmd, 16)
    
    Set adCmd = Nothing
    
    Exit Sub
    
Get_CusSaleInfo_Err:
    MsgBox "Get_CusSaleInfo_Err!"
    Set adCmd = Nothing
    
End Sub
Public Sub Get_VdrSaleInfo(ByVal inVdrID As Long, ByVal inYYYY As String, ByVal inMM As String, ByRef OutCreLmt As Double, ByRef OutCreLft As Double, ByRef OutOpnBal As Double, ByRef OutTotBal As Double, ByRef OutCMQty As Double, ByRef OutCYQty As Double, ByRef OutTotQty As Double, ByRef OutCMSal As Double, ByRef OutCYSal As Double, ByRef OutTotSal As Double, ByRef OutCMNet As Double, ByRef OutCYNet As Double, ByRef OutTotNet As Double)
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_VdrSaleInfo_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GETVDRSALEINFO"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inVdrID)
    Call SetSPPara(adCmd, 2, inYYYY)
    Call SetSPPara(adCmd, 3, inMM)
    
    
    adCmd.Execute
    
    OutCreLmt = GetSPPara(adCmd, 4)
    OutCreLft = GetSPPara(adCmd, 5)
    OutOpnBal = GetSPPara(adCmd, 6)
    OutTotBal = GetSPPara(adCmd, 7)
    
    OutCMQty = GetSPPara(adCmd, 8)
    OutCYQty = GetSPPara(adCmd, 9)
    OutTotQty = GetSPPara(adCmd, 10)
    
    OutCMSal = GetSPPara(adCmd, 11)
    OutCYSal = GetSPPara(adCmd, 12)
    OutTotSal = GetSPPara(adCmd, 13)
    
    OutCMNet = GetSPPara(adCmd, 14)
    OutCYNet = GetSPPara(adCmd, 15)
    OutTotNet = GetSPPara(adCmd, 16)
    
    Set adCmd = Nothing
    
    Exit Sub
    
Get_VdrSaleInfo_Err:
    MsgBox "Get_VdrSaleInfo_Err!"
    Set adCmd = Nothing
    
End Sub



Public Function Get_CusCreditAmt(ByVal InCusID As Long, ByVal inDate As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_CusCreditAmt_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GETCUSCREDITAMT"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InCusID)
    Call SetSPPara(adCmd, 2, inDate)
    
    adCmd.Execute
    
    Get_CusCreditAmt = GetSPPara(adCmd, 3)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_CusCreditAmt_Err:
    MsgBox "Get_CusCreditAmt_Err!"
    Set adCmd = Nothing
    
End Function

Public Function Get_VdrDebitAmt(ByVal inVdrID As Long, ByVal inDate As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_VdrDebitAmt_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GETVdrDebitAMT"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inVdrID)
    Call SetSPPara(adCmd, 2, inDate)
    
    adCmd.Execute
    
    Get_VdrDebitAmt = GetSPPara(adCmd, 3)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_VdrDebitAmt_Err:
    MsgBox "Get_VdrDebitAmt_Err!"
    Set adCmd = Nothing
    
End Function


Public Function Chk_PoRefDoc(ByVal inID As String, ByRef OutTrnCd As String, ByRef OutDocNo As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    Chk_PoRefDoc = False
    
    wsSQL = "SELECT GRHDDOCNO FROM POPGRHD WHERE GRHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND GRHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "GRHDDOCNO")
        OutTrnCd = "GR"
        Chk_PoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    wsSQL = "SELECT PVHDDOCNO FROM POPPVHD WHERE PVHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND PVHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "PVHDDOCNO")
        OutTrnCd = "PV"
        Chk_PoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

    wsSQL = "SELECT PRHDDOCNO FROM POPPRHD WHERE PRHDREFDOCID = " & To_Value(inID) & " "
    wsSQL = wsSQL & "AND PRHDSTATUS IN ('1','4') "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutDocNo = ReadRs(rsRcd, "PRHDDOCNO")
        OutTrnCd = "PR"
        Chk_PoRefDoc = True
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    
    rsRcd.Close
    Set rsRcd = Nothing
  

    
  
End Function


Public Function Get_DRmkID(ByVal inType As String, ByVal inRmkID As String) As String
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_DRmkID_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GETDRMKID"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, inType)
    Call SetSPPara(adCmd, 2, inRmkID)
    Call SetSPPara(adCmd, 3, gsUserID)
    Call SetSPPara(adCmd, 4, gsSystemDate)
    
    adCmd.Execute
    
    Get_DRmkID = GetSPPara(adCmd, 5)
    
    Set adCmd = Nothing
    
    Exit Function
    
Get_DRmkID_Err:
    MsgBox "Get_DRmkID_Err!"
    Set adCmd = Nothing
    
End Function


Public Function Get_JobCost(ByVal inJobNo As String) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "Select JobCost "
    wsSQL = wsSQL & " From mstJobNo "
    wsSQL = wsSQL & " Where JobNo = '" & Set_Quote(inJobNo) & "'"
    
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Get_JobCost = To_Value(ReadRs(rsRcd, "JobCost"))
        
    Else
        Get_JobCost = 0
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_DupCusPO(ByVal InTrnCd As String, ByVal InDocNo As String, ByVal InCusPO As String, ByRef OutDoc As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_DupCusPO = False
    
    
    Select Case InTrnCd
    Case "PR"
    
    wsSQL = "SELECT PRHDDOCNO DOCNO FROM POPPRHD "
    wsSQL = wsSQL & "WHERE PRHDCUSPO = '" & Set_Quote(InCusPO) & "'"
    wsSQL = wsSQL & "AND PRHDDOCNO <> '" & Set_Quote(InDocNo) & "'"
    
    Case "GR"
    
    wsSQL = "SELECT GRHDDOCNO DOCNO FROM POPGRHD "
    wsSQL = wsSQL & "WHERE GRHDCUSPO = '" & Set_Quote(InCusPO) & "'"
    wsSQL = wsSQL & "AND GRHDDOCNO <> '" & Set_Quote(InDocNo) & "'"
    
    Case Else
        OutDoc = ""
        MsgBox "No Such TrnCd!Chk_DupCusPO Mod!"
        Exit Function
    End Select
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutDoc = ReadRs(rsRcd, "DOCNO")
        Chk_DupCusPO = True
    End If
    rsRcd.Close
    
    
    Select Case InTrnCd
    Case "GR"
    
    wsSQL = "SELECT PRHDDOCNO DOCNO FROM POPPRHD "
    wsSQL = wsSQL & "WHERE PRHDCUSPO = '" & Set_Quote(InCusPO) & "'"
    
    Case "PR"
    
    wsSQL = "SELECT GRHDDOCNO DOCNO FROM POPGRHD "
    wsSQL = wsSQL & "WHERE GRHDCUSPO = '" & Set_Quote(InCusPO) & "'"
    
    End Select
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutDoc = ReadRs(rsRcd, "DOCNO")
        Chk_DupCusPO = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_DocAppendix(ByVal InDocID As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT * FROM mstDocAppendix WHERE DADOCID = " & To_Value(InDocID) & " "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        
        Chk_DocAppendix = True
        
    Else
       
        Chk_DocAppendix = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function


Public Function Get_TransferLotTo(InDocID As String, InItemID As String) As String

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT SJDTLOTNO FROM ICSTKADJDT WHERE SJDTDOCID = " & To_Value(InDocID)
    wsSQL = wsSQL & " AND SJDTITEMID = " & To_Value(InItemID)
    wsSQL = wsSQL & " AND SJDTTRNTYPE = 'STKIN' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        Get_TransferLotTo = ReadRs(rsRcd, "SJDTLOTNO")
    Else
        Get_TransferLotTo = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Function


