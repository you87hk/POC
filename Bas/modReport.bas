Attribute VB_Name = "modReport"

Public Function PrintAccessReport(ByVal strReportName As String, ByVal iView As Integer, ByVal sWhereSQL As String) As Boolean


      ' Create a variable that will refer to Microsoft Access.
      Dim MyAccessObject As Access.Application
      Dim bLocal As Boolean
      
      ' Create variables for file pathing.      Dim sPsep, DBPath As String
      ' Set DBPath to the path of Northwind database example.
      ' NOTE: Because Microsoft PowerPoint Visual Basic for
      '       Applications does not recognize the PathSeparator
      '       function, check for the program that is running
      '       this routine. For Microsoft Word and Microsoft Excel,
      '       use the PathSeparator function.

     ' DBPath = Application.Path & sPsep & "Samples" & sPsep _
      & "Northwind.mdb"

     On Error GoTo ErrHandler
      ' Create an instance of the Access application object.
      Set MyAccessObject = CreateObject("Access.Application.8")
      bLocal = False
PrintAgain:          ' < NOTE: This line must be left aligned!
    
      If sLangID = "2" Then
        strReportName = "c" & strReportName
      End If
      
      If iView = 0 Then
       With MyAccessObject
         .OpenCurrentDatabase sDBPath & sDBName  'Opens the Northwind file.
         .DoCmd.OpenReport strReportName, acViewPreview, , sWhereSQL  'Opens the Employees form.
         .Visible = True   ' Maximize Microsoft Access.
       End With
       Else
        With MyAccessObject
         .OpenCurrentDatabase sDBPath & sDBName 'Opens the Northwind file.
         .DoCmd.OpenReport strReportName, acViewNormal, , sWhereSQL 'Opens the Employees form.
         Set MyAccessObject = Nothing ' Set the object variable to
       
       End With
       End If
Exit Function
ErrHandler:          ' < NOTE: This line must be left aligned!
         ' Error handler displays the error message
         ' and then quits the program.
         If Err = 7866 And bLocal = False Then
         sDBPath = "c:\Nbase\Account\"
         bLocal = True
         GoTo PrintAgain
         End If
         
         If Err <> 0 And Err <> 2455 And Err <> 2501 Then
            MsgBox Err & "," & Err.Description, Title:="Test OLE Automation"
            MyAccessObject.Quit   ' Close Microsoft Access.
         PrintAccessReport = False
         ElseIf Err = 2501 Then
            MsgBox GetErrorMessage("RPT0001")
            PrintAccessReport = False
         Else
         PrintAccessReport = True
         End If
         
         Set MyAccessObject = Nothing ' Set the object variable to
                                      ' nothing    End Sub

End Function





Public Function CreateSaleForecast(ByVal sYYYY As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sFromDate As String
    Dim sToDate As String
    Dim sWhereSQL As String
    Dim iTo As Integer
    Dim i, j As Integer
    Dim vMM As Variant
    Dim sToD As String
    
    
    vMM = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", _
                "AUG", "SEP", "OCT", "NOV", "DEC")
On Error GoTo Err_Handler

    

    
    sToD = dtaSystemDate
    sFromDate = Val(sYYYY) - 1 & "/04/01"
    sToDate = sYYYY & "/03/31"
    
    
    sSQL = "Delete From tblRptItemReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    sSQL = "SELECT tblSaleCon.CustomerID, tblItem.Brand, Sum(tblSaleConDetail.AmountInHK) As AmtInHK "
    sSQL = sSQL & "FROM (tblSaleCon INNER JOIN tblSaleConDetail ON tblSaleCon.SaleConID = tblSaleConDetail.SaleConID) INNER JOIN tblItem ON tblSaleConDetail.ItemID = tblItem.ItemID "
    sSQL = sSQL & "WHERE ((tblSaleCon.Status)='4') AND ((tblSaleCon.SaleConDate) >= '" & sFromDate & "') And ((tblSaleCon.SaleConDate) <= '" & sToDate & "') "
    sSQL = sSQL & "GROUP BY tblSaleCon.CustomerID, tblItem.Brand"
    
    
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    rsRcd.Close
    GoTo Err_Handler
    Else
    
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
       
    sSQL = "INSERT INTO tblRptItemReport ( CardID, ReportDate, Brand, YTDAmt, WorkStationID ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("CustomerID") & ", '" & dtaSystemDate & "', '" & rsRcd("Brand") & "', " & rsRcd("AmtInHK") & ", '" & sWorkStationID & "')"
    cnCon.Execute sSQL
    
    rsRcd.MoveNext
    Loop
    
    End If
    rsRcd.Close
    
    
    
    
    For j = 1 To 12
    
    If j <= 9 Then
    i = j + 3
    sFromDate = sYYYY - 1 & "/" & Right(100 + j + 3, 2) & "/01"
    sToDate = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(sFromDate))), "yyyy/mm/dd")
    
    Else
    i = j - 9
    sFromDate = sYYYY & "/" & Right(100 + j - 9, 2) & "/01"
    sToDate = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(sFromDate))), "yyyy/mm/dd")
    
    End If
    
    
    
    
    
    sSQL = "Delete From tblRptItemMTDReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sSQL = "SELECT tblSaleCon.CustomerID, tblItem.Brand, Sum(tblSaleConDetail.Qty) AS TotQty, Sum(tblSaleConDetail.AmountInHK) AS TotAmt "
    sSQL = sSQL & "FROM (tblSaleCon INNER JOIN tblSaleConDetail ON tblSaleCon.SaleConID = tblSaleConDetail.SaleConID) INNER JOIN tblItem ON tblSaleConDetail.ItemID = tblItem.ItemID "
    sSQL = sSQL & "WHERE ((tblSaleCon.Status)='4') AND ((tblSaleCon.SaleConDate) >= '" & sFromDate & "') And ((tblSaleCon.SaleConDate) <= '" & sToDate & "') "
    sSQL = sSQL & "GROUP BY tblSaleCon.CustomerID, tblItem.Brand"
    
    '''
    Set rsRcd = cnCon.Execute(sSQL)
    If Not rsRcd.BOF Or Not rsRcd.EOF Then
       
    sSQL = "INSERT INTO tblRptItemMTDReport ( CardID, ReportDate, Brand, MTDQty, MTDAmt, WorkStationID ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("CustomerID") & ", '" & dtaSystemDate & "', '" & rsRcd("Brand") & "', " & rsRcd("TotQty") & "," & rsRcd("TotAmt") & ",'" & sWorkStationID & "')"
    cnCon.Execute sSQL
    
    sSQL = "UPDATE tblRptItemReport SET tblRptItemReport." & vMM(i - 1) & "Qty = tblRptItemMTDReport.MTDQty, tblRptItemReport." & vMM(i - 1) & "Amt = tblRptItemMTDReport.MTDAmt "
    sSQL = sSQL & "FROM tblRptItemMTDReport,tblRptItemReport "
    sSQL = sSQL & "Where ((tblRptItemMTDReport.Brand = tblRptItemReport.Brand) AND (tblRptItemMTDReport.CardID = tblRptItemReport.CardID) AND ((tblRptItemReport.WorkStationID) = '" & sWorkStationID & "') AND ((tblRptItemMTDReport.WorkStationID) = '" & sWorkStationID & "'))"
    cnCon.Execute sSQL
    
    End If
    rsRcd.Close
    
    
    Next j
    
    
 
    
    
    sSQL = "Select * from tblRptItemReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateSaleForecast = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateSaleForecast = True
Exit Function

Err_Handler:
MsgBox Err.Description
CreateSaleForecast = False

End Function




Public Function CreateARAging(ByVal sCardID As String, ByVal sFromD As String, ByVal sToD As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sWhereSQL As String
    Dim i As Integer
    Dim vMM As Variant
    Dim sInvoiceDate As String
    Dim iDue As String
    vMM = Array("D0d", "D31d", "D61d", "D91d")
                
On Error GoTo Err_Handler

    

 

    sSQL = "Delete From tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sWhereSQL = "Select InvoiceID, InvoiceDate, PromisedDate, CustomerID, BalanceDueInHK "
    sWhereSQL = sWhereSQL & "From tblInvoice "
    sWhereSQL = sWhereSQL & "WHERE (Status='4') AND (BalanceDueInHK > 0) AND (InvoiceDate >= '" & sFromD & "' And InvoiceDate <= '" & sToD & "')"
    
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & " AND (CustomerID=" & sCardID & ")"
    End If
    
    Set rsRcd = cnCon.Execute(sWhereSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    rsRcd.Close
    GoTo Err_Handler
    Else
    
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    
    If Trim(rsRcd("PromisedDate")) <> "" Then
    iDue = DateDiff("d", CDate(dtaSystemDate), CDate(rsRcd("PromisedDate")))
    Else
    iDue = 0
    End If
    
    If iDue > 0 Then
        If iDue < 31 Then
            i = 0
        ElseIf iDue > 30 And iDue < 61 Then
            i = 1
        ElseIf iDue > 60 And iDue < 91 Then
            i = 2
        ElseIf iDue > 90 Then
        i = 3
        End If
    
    sSQL = "INSERT INTO tblRptARReport ( CardID, InvoiceID, ReportDate, FromDate, ToDate, InvoiceDate, PromisedDate, AmtInHK, WorkStationID ," & vMM(i) & "Amt ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("CustomerID") & ", " & rsRcd("InvoiceID") & ", '" & dtaSystemDate & "', '" & sFromD & "', '" & sToD & "', '" & rsRcd("InvoiceDate") & "', '" & rsRcd("PromisedDate") & "', " & rsRcd("BalanceDueInHK") & ", '" & sWorkStationID & "', " & rsRcd("BalanceDueInHK") & ")"
    cnCon.Execute sSQL
    
    End If
    
    rsRcd.MoveNext
    Loop
    
    End If
    rsRcd.Close
    
    
    
    sSQL = "Select * from tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateARAging = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateARAging = True
Exit Function

Err_Handler:
MsgBox Err.Description

CreateARAging = False

End Function



Public Function CreateStatement(ByVal sCardID As String, ByVal sFromD As String, ByVal sToD As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sWhereSQL As String
    Dim i As Integer
    Dim vMM As Variant
    Dim lCardID As Long
    Dim lInvoiceID As Long
    Dim sInvoiceDate As String
                
On Error GoTo Err_Handler

    

    cnCon.BeginTrans

    sSQL = "Delete From tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sSQL = "INSERT INTO tblRptARReport ( ReportDate, FromDate, ToDate, CardID, InvoiceID, InvoiceDate, PromisedDate, AmtInHK, WorkStationID ) "
    sSQL = sSQL & "Select '" & dtaSystemDate & "' As RDate, '" & sFromD & "' As FDate, '" & sToD & "' As TDate, tblInvoice.CustomerID, tblInvoice.InvoiceID, tblInvoice.InvoiceDate, tblInvoice.PromisedDate, tblInvoice.BalanceDueInHK /tblCardFile.Currate As AmtInHK , '" & sWorkStationID & "' As WS "
    sSQL = sSQL & "FROM tblCardFile INNER JOIN tblInvoice ON tblCardFile.CardID = tblInvoice.CustomerID "
    sSQL = sSQL & "WHERE (tblInvoice.Status='4') AND (tblInvoice.BalanceDueInHK > 0) AND (tblInvoice.InvoiceDate >= '" & sFromD & "' And tblInvoice.InvoiceDate <= '" & sToD & "')"
    
    If sCardID <> "0" Then
    sSQL = sSQL & " AND (CustomerID=" & sCardID & ")"
    End If
    
    cnCon.Execute sSQL
    


    cnCon.CommitTrans
    
    
    sSQL = "Select * from tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateStatement = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateStatement = True
Exit Function

Err_Handler:
MsgBox Err.Description
cnCon.RollbackTrans
CreateStatement = False

End Function


Public Function CreateAROverdue(ByVal sCardID As String, ByVal sToD As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sWhereSQL As String
    Dim i As Integer
    Dim vMM As Variant
    Dim sInvoiceDate As String
    Dim iDue As String
    vMM = Array("D0d", "D31d", "D61d", "D91d")
                
On Error GoTo Err_Handler

    

  
    sSQL = "Delete From tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sWhereSQL = "Select InvoiceID, InvoiceDate, PromisedDate, CustomerID, BalanceDueInHK "
    sWhereSQL = sWhereSQL & "From tblInvoice "
    sWhereSQL = sWhereSQL & "WHERE (Status='4') AND (BalanceDueInHK > 0) AND (PromisedDate < '" & sToD & "')"
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & " AND (CustomerID=" & sCardID & ")"
    End If
    Set rsRcd = cnCon.Execute(sWhereSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    rsRcd.Close
    GoTo Err_Handler
    Else
    
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    
      
    If Trim(rsRcd("PromisedDate")) <> "" Then
    iDue = DateDiff("d", CDate(rsRcd("PromisedDate")), CDate(sToD))
    Else
    iDue = 0
    End If
    
    If iDue > 0 Then
        If iDue < 31 Then
            i = 0
        ElseIf iDue > 30 And iDue < 61 Then
            i = 1
        ElseIf iDue > 60 And iDue < 91 Then
            i = 2
        ElseIf iDue > 90 Then
        i = 3
        End If
    
    sSQL = "INSERT INTO tblRptARReport ( CardID, InvoiceID, ReportDate, InvoiceDate, PromisedDate, AmtInHK, WorkStationID ," & vMM(i) & "Amt ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("CustomerID") & ", " & rsRcd("InvoiceID") & ", '" & sToD & "', '" & rsRcd("InvoiceDate") & "', '" & rsRcd("PromisedDate") & "', " & rsRcd("BalanceDueInHK") & ", '" & sWorkStationID & "', " & rsRcd("BalanceDueInHK") & ")"
    cnCon.Execute sSQL
    
    End If
    
    rsRcd.MoveNext
    Loop
    
    End If
    rsRcd.Close
    

    
    
    sSQL = "Select * from tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateAROverdue = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateAROverdue = True
Exit Function

Err_Handler:
MsgBox Err.Description

CreateAROverdue = False

End Function


Public Function CreateAPAging(ByVal sCardID As String, ByVal sFromD As String, ByVal sToD As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sWhereSQL As String
    Dim i As Integer
    Dim vMM As Variant
    Dim sInvoiceDate As String
    Dim iDue As String
    vMM = Array("D0d", "D31d", "D61d", "D91d")
                
On Error GoTo Err_Handler

    

   
    sSQL = "Delete From tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sWhereSQL = "Select PurchaseOrderID, PurchaseOrderDate, PromisedDate, SupplierID, BalanceDueInHK "
    sWhereSQL = sWhereSQL & "From tblPurchaseOrder "
    sWhereSQL = sWhereSQL & "WHERE (Status='4') AND (BalanceDueInHK > 0) AND (PurchaseOrderDate >= '" & sFromD & "' And PurchaseOrderDate <= '" & sToD & "')"
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & " AND (SupplierID=" & sCardID & ")"
    End If
    Set rsRcd = cnCon.Execute(sWhereSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    rsRcd.Close
    GoTo Err_Handler
    Else
    
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    If Trim(rsRcd("PromisedDate")) <> "" Then
    iDue = DateDiff("d", CDate(dtaSystemDate), CDate(rsRcd("PromisedDate")))
    Else
    iDue = 0
    End If
    
    If iDue > 0 Then
        If iDue < 31 Then
            i = 0
        ElseIf iDue > 30 And iDue < 61 Then
            i = 1
        ElseIf iDue > 60 And iDue < 91 Then
            i = 2
        ElseIf iDue > 90 Then
        i = 3
        End If
    
    sSQL = "INSERT INTO tblRptARReport ( CardID, InvoiceID, ReportDate, FromDate, ToDate, InvoiceDate, PromisedDate, AmtInHK, WorkStationID ," & vMM(i) & "Amt ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("SupplierID") & ", " & rsRcd("PurchaseOrderID") & ", '" & dtaSystemDate & "', '" & sFromD & "', '" & sToD & "', '" & rsRcd("PurchaseOrderDate") & "', '" & rsRcd("PromisedDate") & "', " & rsRcd("BalanceDueInHK") & ", '" & sWorkStationID & "', " & rsRcd("BalanceDueInHK") & ")"
    cnCon.Execute sSQL
    
    End If
    
    rsRcd.MoveNext
    Loop
    
    End If
    rsRcd.Close
    

    
    
    sSQL = "Select * from tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateAPAging = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateAPAging = True
Exit Function

Err_Handler:
MsgBox Err.Description

CreateAPAging = False

End Function
Public Function CreateAPOverdue(ByVal sCardID As String, ByVal sToD As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sWhereSQL As String
    Dim i As Integer
    Dim vMM As Variant
    Dim sInvoiceDate As String
    Dim iDue As String
    vMM = Array("D0d", "D31d", "D61d", "D91d")
                
On Error GoTo Err_Handler

    

 

    sSQL = "Delete From tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    
    sWhereSQL = "Select PurchaseOrderID, PurchaseOrderDate, PromisedDate, SupplierID, BalanceDueInHK "
    sWhereSQL = sWhereSQL & "From tblPurchaseOrder "
    sWhereSQL = sWhereSQL & "WHERE (Status='4') AND (BalanceDueInHK > 0) AND (PromisedDate < '" & sToD & "')"
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & " AND (SupplierID=" & sCardID & ")"
    End If
    Set rsRcd = cnCon.Execute(sWhereSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    rsRcd.Close
    GoTo Err_Handler
    Else
    
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    
    
    If Trim(rsRcd("PromisedDate")) <> "" Then
    iDue = DateDiff("d", CDate(rsRcd("PromisedDate")), CDate(sToD))
    Else
    iDue = 0
    End If
    
    
    If iDue > 0 Then
        If iDue < 31 Then
            i = 0
        ElseIf iDue > 30 And iDue < 61 Then
            i = 1
        ElseIf iDue > 60 And iDue < 91 Then
            i = 2
        ElseIf iDue > 90 Then
        i = 3
        End If
    
    sSQL = "INSERT INTO tblRptARReport ( CardID, InvoiceID, ReportDate, InvoiceDate, PromisedDate, AmtInHK, WorkStationID ," & vMM(i) & "Amt ) "
    sSQL = sSQL & "VALUES ( " & rsRcd("SupplierID") & ", " & rsRcd("PurchaseOrderID") & ", '" & sToD & "', '" & rsRcd("PurchaseOrderDate") & "', '" & rsRcd("PromisedDate") & "', " & rsRcd("BalanceDueInHK") & ", '" & sWorkStationID & "', " & rsRcd("BalanceDueInHK") & ")"
    cnCon.Execute sSQL
    
    End If
    
    rsRcd.MoveNext
    Loop
    
    End If
    rsRcd.Close
    
    cnCon.CommitTrans
    
    
    sSQL = "Select * from tblRptARReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateAPOverdue = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateAPOverdue = True
Exit Function

Err_Handler:
MsgBox Err.Description

CreateAPOverdue = False

End Function
Public Function CreateOSOrder(ByVal sCardID As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sFromDate As String
    Dim sWhereSQL As String
    
   
On Error GoTo Err_Handler

    

    cnCon.BeginTrans

    sWhereSQL = "WHERE ( tblSaleCon.Status='4' AND tblSaleConDetail.RemainQty > 0 ) "
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & "AND ( tblSaleCon.CustomerID=" & sCardID & ") "
    End If
    
    sSQL = "Delete From tblRptOSReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    sSQL = "INSERT INTO tblRptOSReport ( ReportDate, TranID, WorkStationID ) "
    sSQL = sSQL & "SELECT '" & dtaSystemDate & "' AS RDate, tblSaleConDetail.SaleConDTLID, '" & sWorkStationID & "' As Wsts "
    sSQL = sSQL & "FROM tblSaleCon INNER JOIN tblSaleConDetail ON tblSaleCon.SaleConID = tblSaleConDetail.SaleConID "
    sSQL = sSQL & sWhereSQL
    cnCon.Execute sSQL
    
   
    cnCon.CommitTrans
    
    
    sSQL = "Select * from tblRptOSReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateOSOrder = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateOSOrder = True
Exit Function

Err_Handler:
MsgBox Err.Description
cnCon.RollbackTrans
CreateOSOrder = False

End Function
Public Function CreateOSPO(ByVal sCardID As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim sFromDate As String
    Dim sWhereSQL As String
    
   
On Error GoTo Err_Handler

    

    cnCon.BeginTrans

    sWhereSQL = "WHERE ( tblPurchaseOrder.Status='4' AND tblPurchaseOrderDetail.RemainQty > 0 ) "
    If sCardID <> "0" Then
    sWhereSQL = sWhereSQL & "AND ( tblPurchaseOrder.SupplierID=" & sCardID & ") "
    End If
    
    sSQL = "Delete From tblRptOSReport Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
    
    sSQL = "INSERT INTO tblRptOSReport ( ReportDate, TranID, WorkStationID ) "
    sSQL = sSQL & "SELECT '" & dtaSystemDate & "' AS RDate, tblPurchaseOrderDetail.PurchaseOrderDTLID, '" & sWorkStationID & "' As Wsts "
    sSQL = sSQL & "FROM tblPurchaseOrder INNER JOIN tblPurchaseOrderDetail ON tblPurchaseOrder.PurchaseOrderID = tblPurchaseOrderDetail.PurchaseOrderID "
    sSQL = sSQL & sWhereSQL
    cnCon.Execute sSQL
    
   
    cnCon.CommitTrans
    
    
    sSQL = "Select * from tblRptOSReport Where WorkStationID = '" & sWorkStationID & "'"
    Set rsRcd = cnCon.Execute(sSQL)
    If rsRcd.BOF Or rsRcd.EOF Then
    MsgBox GetErrorMessage("RPT0001")
    CreateOSPO = False
    rsRcd.Close
    Exit Function
    End If
    rsRcd.Close
    
Quit_Func:
CreateOSPO = True
Exit Function

Err_Handler:
MsgBox Err.Description
cnCon.RollbackTrans
CreateOSPO = False

End Function

Public Function CreateSaleTarget(ByVal FDate As String, ByVal TDate As String, ByVal sCardID As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim i, j As Integer
    Dim vMM As Variant
    Dim iFYYYY As Long
    Dim iTYYYY As Long
    Dim iFMM As Integer
    Dim iTMM As Integer
    Dim iYYYY As Long
    Dim iMM As Integer
    
    Dim iYYYYMM As Long
    Dim iTYYYYMM As Long
    Dim dAmt As Currency
    
    iFYYYY = Left(FDate, 4)
    iTYYYY = Left(TDate, 4)
    iFMM = Mid(FDate, 6, 2)
    iTMM = Mid(TDate, 6, 2)
    
    
    vMM = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", _
                "AUG", "SEP", "OCT", "NOV", "DEC")
                
On Error GoTo Err_Handler

    iYYYY = iFYYYY
    iMM = iFMM
    
    iTYYYYMM = iTYYYY * 100 + iTMM
    iYYYYMM = iYYYY * 100 + iMM
    cnCon.BeginTrans
    
       
    sSQL = "Delete From tblRptSaleTar Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
 
    
    Do Until (iYYYYMM > iTYYYYMM)
    
    sFromDate = iYYYY & "/" & Right(100 + iMM, 2) & "/01"
    sToDate = iYYYY & "/" & Right(100 + iMM, 2) & "/31"
    
    
    sSQL = "SELECT Sum(tblSaleCon.TotalAmountInHK) AS TotAmtInHK "
    sSQL = sSQL & "From tblSaleCon "
    sSQL = sSQL & "Where (((tblSaleCon.SalesmanID) = " & sCardID & ") And ((tblSaleCon.SaleConDate) >= '" & sFromDate & "' And (tblSaleCon.SaleConDate) <= '" & sToDate & "') And ((tblSaleCon.Status) = '4'))"
    
    
    Set rsRcd = cnCon.Execute(sSQL)
    If IsNull(rsRcd("TotAmtInHK")) Then
        dAmt = 0
    Else
        dAmt = rsRcd("TotAmtInHK")
    End If
    rsRcd.Close
    
       
    sSQL = "INSERT INTO tblRptSaleTar ( YYYY, MonthID, YYYYMM, CardID, Target, SalesVolume, WorkStationID ) "
    sSQL = sSQL & "VALUES ( '" & iYYYY & "', " & iMM & ", '" & vMM(iMM - 1) & "-" & Right(iYYYY, 2) & "', " & sCardID & ", 0, " & dAmt & ", '" & sWorkStationID & "')"
    cnCon.Execute sSQL
    
    
    iMM = iMM + 1
    
    If iMM = 13 Then
    iYYYY = iYYYY + 1
    iMM = 1
    End If
    
    iYYYYMM = iYYYY * 100 + iMM
    
    Loop
    
    
    sSQL = "UPDATE tblRptSaleTar SET tblRptSaleTar.Target = tblSaleTar.Target "
    sSQL = sSQL & "From tblRptSaleTar, tblSaleTar "
    sSQL = sSQL & "Where (((tblRptSaleTar.WorkStationID) = '" & sWorkStationID & "') And (tblRptSaleTar.CardID = tblSaleTar.SalesmanID) And (tblRptSaleTar.YYYY = tblSaleTar.YYYY))"
    cnCon.Execute sSQL
    
    cnCon.CommitTrans
    
    
Quit_Func:
CreateSaleTarget = True
Exit Function

Err_Handler:
cnCon.RollbackTrans
MsgBox Err.Description
CreateSaleTarget = False

End Function

Public Function CreatePurTarget(ByVal FDate As String, ByVal TDate As String, ByVal sCardID As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    Dim i, j As Integer
    Dim vMM As Variant
    Dim iFYYYY As Long
    Dim iTYYYY As Long
    Dim iFMM As Integer
    Dim iTMM As Integer
    Dim iYYYY As Long
    Dim iMM As Integer
    
    Dim iYYYYMM As Long
    Dim iTYYYYMM As Long
    Dim dAmt As Currency
    
    iFYYYY = Left(FDate, 4)
    iTYYYY = Left(TDate, 4)
    iFMM = Mid(FDate, 6, 2)
    iTMM = Mid(TDate, 6, 2)
    
    
    vMM = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", _
                "AUG", "SEP", "OCT", "NOV", "DEC")
                
On Error GoTo Err_Handler

    iYYYY = iFYYYY
    iMM = iFMM
    
    iTYYYYMM = iTYYYY * 100 + iTMM
    iYYYYMM = iYYYY * 100 + iMM
    cnCon.BeginTrans
    
       
    sSQL = "Delete From tblRptSaleTar Where WorkStationID = '" & sWorkStationID & "'"
    cnCon.Execute sSQL
 
    
    Do Until (iYYYYMM > iTYYYYMM)
    
    sFromDate = iYYYY & "/" & Right(100 + iMM, 2) & "/01"
    sToDate = iYYYY & "/" & Right(100 + iMM, 2) & "/31"
    
    
    sSQL = "SELECT Sum(tblPurchaseOrder.TotalAmountInHK) AS TotAmtInHK "
    sSQL = sSQL & "From tblPurchaseOrder "
    sSQL = sSQL & "Where (((tblPurchaseOrder.SupplierID) = " & sCardID & ") And ((tblPurchaseOrder.PurchaseOrderDate) >= '" & sFromDate & "' And (tblPurchaseOrder.PurchaseOrderDate) <= '" & sToDate & "') And ((tblPurchaseOrder.Status) = '4'))"
    
    
    Set rsRcd = cnCon.Execute(sSQL)
    If IsNull(rsRcd("TotAmtInHK")) Then
        dAmt = 0
    Else
        dAmt = rsRcd("TotAmtInHK")
    End If
    rsRcd.Close
    
       
    sSQL = "INSERT INTO tblRptSaleTar ( YYYY, MonthID, YYYYMM, CardID, Target, SalesVolume, WorkStationID ) "
    sSQL = sSQL & "VALUES ( '" & iYYYY & "', " & iMM & ", '" & vMM(iMM - 1) & "-" & Right(iYYYY, 2) & "', " & sCardID & ", 0, " & dAmt & ", '" & sWorkStationID & "')"
    cnCon.Execute sSQL
    
    
    iMM = iMM + 1
    
    If iMM = 13 Then
    iYYYY = iYYYY + 1
    iMM = 1
    End If
    
    iYYYYMM = iYYYY * 100 + iMM
    
    Loop
    
    
    sSQL = "UPDATE tblRptSaleTar SET tblRptSaleTar.Target = tblSaleTar.Target "
    sSQL = sSQL & "From tblRptSaleTar, tblSaleTar "
    sSQL = sSQL & "Where (((tblRptSaleTar.WorkStationID) = '" & sWorkStationID & "') And (tblRptSaleTar.CardID = tblSaleTar.SalesmanID) And (tblRptSaleTar.YYYY = tblSaleTar.YYYY))"
    cnCon.Execute sSQL
    
    cnCon.CommitTrans
    
    
Quit_Func:
CreatePurTarget = True
Exit Function

Err_Handler:
cnCon.RollbackTrans
MsgBox Err.Description
CreatePurTarget = False

End Function
