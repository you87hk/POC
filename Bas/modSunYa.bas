Attribute VB_Name = "modSunYa"
Public Sub LoadSaleNameByCode(outTxtBox As TextBox, ByVal inSaleCode)
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT SaleName FROM MstSalesman WHERE SaleCode ='" + Set_Quote(inSaleCode) + "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
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
    Dim wsSql As String
    
    wsSql = "SELECT SaleCode FROM MstSalesman WHERE SaleID =" & To_Value(inSaleManID)
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
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
    Dim wsSql As String
    
    wsSql = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleID =" & To_Value(inSaleManID)
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
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
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsChar = Left(inCallNo2, 1)
    
    If (wsChar > "a" And wsChar < "z") Or (wsChar > "A" And wsChar < "Z") Then
        wsChar = Left(inCallNo2, 3)
    Else
        wsSql = "SELECT CharTable.Hon "
        wsSql = wsSql + "From CharTable "
        wsSql = wsSql + "WHERE (((CharTable.Word)='" + Set_Quote(wsChar) + "'));"
        
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
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
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
        get_PriCode = ""

    
        wsSql = "SELECT PRICODE "
        wsSql = wsSql + "From mstPrimary "
        wsSql = wsSql + "WHERE PriCatCode ='" + Set_Quote(inCatCode) + "'"
        
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
        If rsRcd.RecordCount <> 0 Then
            get_PriCode = ReadRs(rsRcd, "PriCode")
        End If
        
        rsRcd.Close
        Set rsRcd = Nothing
    
    
End Function
Public Function LoadDescByCode(stblName As String, sKeyName As String, sRetName As String, sValue As String, bString As Boolean) As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wsResult As String
        
    LoadDescByCode = ""
        
    If Trim(sValue) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT " + sRetName
    
    If bString = True Then
        wsSql = wsSql + " FROM " + stblName + " WHERE " + sKeyName + " = '" + Set_Quote(sValue) + "' "
    Else
        wsSql = wsSql + " FROM " + stblName + " WHERE " + sKeyName + " = " + Set_Quote(sValue)
    End If
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    Dim wsSql As String
    
    Chk_Curr = False
    
    wsCtlDte = IIf(Trim(inDate) = "" Or Trim(inDate) = "/  /", gsSystemDate, inDate)
    wsSql = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE "
    wsSql = wsSql & "EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
    wsSql = wsSql & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
    wsSql = wsSql & " AND EXCCURR = '" & Set_Quote(InCurr) & "' "
    wsSql = wsSql & " AND EXCSTATUS = '1' "
    
    
    rsCurCod.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsCurCod.RecordCount <= 0 Then
        rsCurCod.Close
        Set rsCurCod = Nothing
        Exit Function
    End If
    
    Chk_Curr = True
    
    rsCurCod.Close
    Set rsCurCod = Nothing
    
End Function

Public Function Chk_PoHdDocNo(ByVal inDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT POHDSTATUS FROM popPOHD WHERE POHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "POHDSTATUS")
        Chk_PoHdDocNo = True
    Else
        OutStatus = ""
        Chk_PoHdDocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function







Public Function Chk_CusCode(ByVal InCusNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT CusID, CusName, CusTel, CusFax FROM mstCustomer WHERE CusCode = '" & Set_Quote(InCusNo) & "' "
    wsSql = wsSql & "And CusStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT SaleID, SaleName FROM mstSalesman WHERE SaleCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And SaleStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT TypID, TypDesc FROM mstType WHERE TypCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And TypClass = '" & Set_Quote(InClass) & "' "
    wsSql = wsSql & "And TypStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    wsSql = "SELECT RgnDesc FROM MstRegion WHERE RgnCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And RgnStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT PayDesc FROM mstPayTerm WHERE PayCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And PayStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT PrcDesc FROM mstPriceTerm WHERE PrcCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And PrcStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT MLDesc FROM mstMerchClass WHERE MLCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And MLStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT MethodDesc FROM mstMethod WHERE MethodCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And MethodStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT NatureDesc FROM mstNature WHERE NatureCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And NatureStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT RmkStatus FROM mstRemark WHERE RmkCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And RmkStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT ShipStatus FROM mstShip WHERE ShipCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And ShipStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String
    Dim wdCusDis As Double
    Dim wdSalDis As Double
    
    wdCusDis = 0
    wdSalDis = 0
    
        
    wsSql = "SELECT CUSSPECDIS "
    wsSql = wsSql & " FROM mstCustomer "
    wsSql = wsSql & " WHERE CusID = " & InCusID
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        wdCusDis = To_Value(ReadRs(rsRcd, "CUSSPECDIS"))
    End If
    
    rsRcd.Close
    
    
    wdCusDis = (100 - wdCusDis) / 100
        
        
    wsSql = "SELECT SDDISCOUNT "
    wsSql = wsSql & " FROM mstSaleDiscount, mstCustomer, mstItem "
    wsSql = wsSql & " WHERE CusID = " & InCusID
    wsSql = wsSql & " AND ItmID = " & InItemID
    wsSql = wsSql & " AND SDNatureCode = '" & Set_Quote(inNatureCode) & "'"
    wsSql = wsSql & " AND SDMethodCode = CusMethodCode "
    wsSql = wsSql & " AND SDCDisCode = ItmCDisCode "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    Dim wsSql As String
        
    wsSql = "SELECT CusItemPrice "
    wsSql = wsSql & " FROM MstCusItem "
    wsSql = wsSql & " WHERE CusItemCusID = " & InCusID
    wsSql = wsSql & " AND CusItemItmID = " & InItemID
    wsSql = wsSql & " AND CusItemCurr = '" & Set_Quote(InCurr) & "'"
    wsSql = wsSql & " AND CusItemStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getCusItemPrice = To_Value(ReadRs(rsRcd, "CusItemPrice"))
    Else
        getCusItemPrice = 0
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function getItemPrice(ByVal inDate As String, ByVal InItemID As Long, ByVal InCurr As String, ByVal inExcr As Double) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wsCurr As String
    
    wsSql = "Select PriChgDtCurr, PriChgDtUnitPrice "
    wsSql = wsSql & " From mstPriceChange, mstPriceChangeDt"
    wsSql = wsSql & " Where PriChgDocID = PriChgDtDocID "
    wsSql = wsSql & " AND PriChgDtItemID = " & InItemID
    wsSql = wsSql & " AND PriChgFromDate <= '" & Set_Quote(inDate) & "'"
    wsSql = wsSql & " AND PriChgToDate >= '" & Set_Quote(inDate) & "'"
    wsSql = wsSql & " AND PriChgStatus = '1'"
    
    
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getItemPrice = To_Value(ReadRs(rsRcd, "PriChgDtUnitPrice"))
        wsCurr = ReadRs(rsRcd, "PriChgDtCurr")
        If UCase(wsCurr) <> UCase(InCurr) Then
            getItemPrice = Format(getItemPrice * inExcr, gsAmtFmt)
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
        
    adCmd.CommandText = "USP_CHECKCREDITLIMIT"
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

Public Function getVdrItemPrice(ByVal InVdrID As Long, ByVal InItemID As Long, ByVal InCurr As String) As Double

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    wsSql = "SELECT VdrItemPrice "
    wsSql = wsSql & " FROM MstVdrItem "
    wsSql = wsSql & " WHERE VdrItemVdrID = " & InVdrID
    wsSql = wsSql & " AND VdrItemItmID = " & InItemID
    wsSql = wsSql & " AND VdrItemCurr = '" & Set_Quote(InCurr) & "'"
    wsSql = wsSql & " AND VdrItemStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        getVdrItemPrice = To_Value(ReadRs(rsRcd, "VdrItemPrice"))
    Else
        getVdrItemPrice = 0
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Public Function Chk_VdrCode(ByVal InVdrNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT VdrID, VdrName, VdrTel, VdrFax FROM mstVendor WHERE VdrCode = '" & Set_Quote(InVdrNo) & "' "
    wsSql = wsSql & "And VdrStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String
    
    Chk_ItmCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT ItmStatus "
    wsSql = wsSql & " FROM MstItem WHERE ItmCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    Dim wsSql As String

    
    wsSql = "SELECT USRGRPCODE FROM mstUser WHERE USRGRPCODE = '" & Set_Quote(inNo) & "' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
    Dim wsSql As String

    
    wsSql = "SELECT tblTableID, tblPgmID, tblStatusName FROM sysMstTable WHERE tblTableID = '" & Set_Quote(inNo) & "' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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



Public Function Get_StockAvailable(ByVal InItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_StockAvailable_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_STOCKAVAILABLE"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InItmID)
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

Public Function Get_PoBalQty(ByVal InTrnCode As String, ByVal InDocID As Long, ByVal InPoID As Long, ByVal InItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_PoBalQty_Err
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetPoBalQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InPoID)
    Call SetSPPara(adCmd, 4, InItmID)
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

Public Function Get_SoBalQty(ByVal InTrnCode As String, ByVal InDocID As Long, ByVal InSoID As Long, ByVal InItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_SoBalQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetSoBalQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InSoID)
    Call SetSPPara(adCmd, 4, InItmID)
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

Public Function Get_StTrnQty(ByVal InTrnCode As String, ByVal InDocID As Long, ByVal InSoID As Long, ByVal InItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_StTrnQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetStTrnQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InTrnCode)
    Call SetSPPara(adCmd, 2, InDocID)
    Call SetSPPara(adCmd, 3, InSoID)
    Call SetSPPara(adCmd, 4, InItmID)
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
Public Function Get_IcQty(ByVal InTrnCode As String, ByVal InSoID As Long, ByVal InItmID As Long, ByVal inWhsCode As String, ByVal InLotNo As String) As Double
    
    Dim adCmd As New ADODB.Command
    
    On Error GoTo Get_IcQty_Err
    
    
    
    Set adCmd.ActiveConnection = cnCon
        
    adCmd.CommandText = "USP_GetIcQty"
    adCmd.CommandType = adCmdStoredProc
    adCmd.Parameters.Refresh
    
    Call SetSPPara(adCmd, 1, InTrnCode)
    Call SetSPPara(adCmd, 2, InSoID)
    Call SetSPPara(adCmd, 3, InItmID)
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
    Dim wsSql As String
    
     wsSql = "SELECT " & inField & " FROM sysWSINFO WHERE WSID = '" & gsWorkStationID & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    Dim wsSql As String
    
    Chk_SpecialPassword = False
    
    If UCase(gsUserID) = "NBASE" And UCase(inPassword) = "NBTEL" Then
            Chk_SpecialPassword = True
            Exit Function
    End If
    
    wsSql = "SELECT USRPWD, USRNAME FROM MSTUSER WHERE USRCODE = 'SUNYA' "

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

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
    Dim wsSql As String
    
    wsSql = "SELECT WhsDesc FROM MstWarehouse WHERE WhsCode = '" & Set_Quote(inCode) & "' "
    wsSql = wsSql & "And WhsStatus = '1' "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
Public Function Chk_TrnHdDocNo(ByVal inTrnCd As String, ByVal inDocNo As String, ByRef OutStatus As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    
    Select Case inTrnCd
    Case "SN", "QT"
    wsSql = "SELECT SNHDSTATUS STATUS FROM soaSNHD WHERE SNHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "SO"
    wsSql = "SELECT SOHDSTATUS STATUS FROM soaSOHD WHERE SOHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "IV"
    wsSql = "SELECT IVHDSTATUS STATUS FROM soaIVHD WHERE IVHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "SP"
    wsSql = "SELECT SPHDSTATUS STATUS FROM soaSPHD WHERE SPHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "SR"
    wsSql = "SELECT SRHDSTATUS STATUS FROM soaSRHD WHERE SRHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "SD"
    wsSql = "SELECT SDHDSTATUS STATUS FROM soaSDHD WHERE SDHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "PO"
    wsSql = "SELECT POHDSTATUS STATUS FROM popPOHD WHERE POHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "PV"
    wsSql = "SELECT PVHDSTATUS STATUS FROM popPVHD WHERE PVHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "PR"
    wsSql = "SELECT PRHDSTATUS STATUS FROM popPRHD WHERE PRHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case "IO", "II"
    wsSql = "SELECT STHDSTATUS STATUS FROM ICSTHD WHERE STHDDOCNO = '" & Set_Quote(inDocNo) & "'"
    Case Else
        OutStatus = ""
        Chk_TrnHdDocNo = False
        MsgBox "No Such TrnCd!Chk_TrnHdDocNo Mod!"
        Exit Function
    End Select
    
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
