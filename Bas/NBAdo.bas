Attribute VB_Name = "NBAdo"

Public Function Connect_Database() As Integer

    Connect_Database = True

    On Error GoTo Err_Handler

    With cnCon
        .Provider = "SQLOLEDB"
        'Modified by Lewis at 09152002
         '.ConnectionTimeout = 10
         .ConnectionTimeout = giTimeOut
         .CommandTimeout = giTimeOut
        .CursorLocation = adUseClient
        .ConnectionString = gsConnectString
        .Open
    End With

    Exit Function
    
Err_Handler:
    Connect_Database = False
    MsgBox "Err Connecting Database! " & Err.Description & " " & Err.Number

End Function
Public Sub Disconnect_Database()

    cnCon.Close
    
    Set cnCon = Nothing
    
End Sub

Public Function ReadRs(ByVal inRs As ADODB.Recordset, inCol As Variant) As Variant
    
    'inCol is the column no (0 based) or column name
    
    Dim TmpCol As Variant
    
    On Error GoTo ReadRs_Err
    
    TmpCol = inCol
    
    If inRs Is Nothing Then Exit Function
    
    If inRs.RecordCount <= 0 Then Exit Function
    
    If IsNumeric(TmpCol) Then
        TmpCol = inCol
        Select Case TmpCol
            Case Is < 0:
                TmpCol = 0
        End Select
    
    
    End If
    
    ReadRs = inRs(TmpCol).Value
    
    If IsNull(ReadRs) Then
        Select Case inRs.Fields(TmpCol).Type
            Case adChar, adVarChar, adVarWChar, adWChar
                ReadRs = ""
            Case adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
                ReadRs = "0"
            Case Else
                ReadRs = ""
        End Select
    End If
    
    Exit Function
    
ReadRs_Err:
    Exit Function

End Function
Public Function SetSPPara(incmd As ADODB.Command, ByVal inPara As Integer, ByVal InValue As Variant)

    If IsEmpty(InValue) Or IsNull(InValue) Then
        InValue = ""
    End If
    
    With incmd.Parameters(inPara)
        Select Case .Type
        Case adChar, adVarChar, adVarWChar, adWChar
            If Len(InValue) > .Size Then
            InValue = Left(InValue, .Size)
            End If
        Case adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
            If IsNumeric(InValue) Then
                InValue = To_Value(InValue)
            Else
                InValue = 0
            End If
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            If InValue = "" Then
               InValue = Null
            End If
        End Select
        
    End With
    
    incmd.Parameters(inPara).Value = InValue

End Function

Public Function GetSPPara(incmd As ADODB.Command, ByVal inPara As Integer) As Variant

    GetSPPara = incmd.Parameters(inPara).Value
    
End Function

